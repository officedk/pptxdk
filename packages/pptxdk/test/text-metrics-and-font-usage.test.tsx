import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { readFileSync } from "node:fs";

import {
  BuildError,
  build,
  defineSlideLayout,
  emit,
  font,
  resolve,
  type FontUsage,
  type Rect,
  type TextMetrics,
} from "@officedk/pptxdk";
import { preview } from "@officedk/pptxdk/preview";
import {
  Box,
  Cell,
  Chart,
  Deck,
  Group,
  Notes,
  P,
  Placeholder,
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
import { createElement } from "react";
import { describe, it } from "vitest";

const jostBookBytes = new Uint8Array(
  readFileSync(new URL("./assets/fonts/Jost-400-Book.ttf", import.meta.url)),
);
const jostBoldBytes = new Uint8Array(
  readFileSync(new URL("./assets/fonts/Jost-700-Bold.ttf", import.meta.url)),
);

const jostFont = font("Jost*", [
  { weight: 400, data: jostBookBytes },
  { weight: 700, data: jostBoldBytes },
]);

const jostRegularOnly = font("Jost*", [{ weight: 400, data: jostBookBytes }]);

function withFontFsType(bytes: Uint8Array, fsType: number): Uint8Array {
  const copy = new Uint8Array(bytes);
  const view = new DataView(copy.buffer, copy.byteOffset, copy.byteLength);
  const tableCount = view.getUint16(4);
  for (let index = 0; index < tableCount; index += 1) {
    const record = 12 + index * 16;
    const tag = String.fromCharCode(...copy.subarray(record, record + 4));
    if (tag === "OS/2") {
      view.setUint16(view.getUint32(record + 8) + 8, fsType);
      return copy;
    }
  }
  throw new Error("Font fixture must contain an OS/2 table");
}

// Pre-change SHA-256 hashes captured from table.ts warning-driven <clipPath> baseline.
const EXPECTED_OVERFLOW_PNG_SHA256 =
  "4c572e731a9b62266a9fd001824999887f01caa335ee0ae280b3ac7d04f1e707";
const EXPECTED_CONTROL_PNG_SHA256 =
  "83609a9d0690e42b476ea241691e85c31725b5957cba460a66e19f10961649ff";

function assertMeasured(
  metrics: TextMetrics | undefined,
): asserts metrics is Extract<TextMetrics, { readonly status: "measured" }> {
  assert.ok(metrics !== undefined, "expected textMetrics to be defined");
  assert.equal(metrics.status, "measured");
}

function assertUnmeasured(
  metrics: TextMetrics | undefined,
): asserts metrics is Extract<TextMetrics, { readonly status: "unmeasured" }> {
  assert.ok(metrics !== undefined, "expected textMetrics to be defined");
  assert.equal(metrics.status, "unmeasured");
}

describe("Ticket 01 Slice 1: Rect, tri-state TextMetrics & warning-free table cell clipping", () => {
  it("computes measured TextMetrics with padding, preset textRect, fits tolerance, and unmeasured reasons", async () => {
    const layoutDef = defineSlideLayout({
      id: "title-layout",
      placeholders: {
        title: {
          type: "title" as const,
          x: 40,
          y: 20,
          width: 400,
          height: 80,
        },
      },
    });
    const doc = createElement(
      Deck,
      { width: 960, height: 540 },
      createElement(
        SlideMaster,
        { id: "master" },
        createElement(SlideLayout, { layout: layoutDef }),
      ),
      createElement(
        Slide,
        {
          id: "metrics-slide",
          layout: layoutDef,
          dateTime: "2026-11-25",
          footer: "Confidential",
          slideNumber: true,
        },
        createElement(
          Text,
          {
            id: "compact-overflow",
            x: 40,
            y: 40,
            width: 460,
            height: 110,
            style: { fontFamily: jostRegularOnly, fontSize: 100 },
          },
          "2025",
        ),
        createElement(
          Text,
          {
            id: "compact-fits",
            x: 40,
            y: 160,
            width: 460,
            height: 150,
            style: { fontFamily: jostRegularOnly, fontSize: 100 },
          },
          "2025",
        ),
        createElement(
          Text,
          {
            id: "auto-text",
            x: 40,
            y: 320,
            width: 240,
            style: { fontFamily: jostRegularOnly, fontSize: 20, padding: 8 },
          },
          "Auto sized line",
        ),
        createElement(
          Text,
          {
            id: "empty-fixed",
            x: 300,
            y: 320,
            width: 200,
            height: 40,
            style: { fontFamily: jostRegularOnly, fontSize: 16 },
          },
          "",
        ),
        createElement(
          Text,
          {
            id: "missing-font-fixed",
            x: 520,
            y: 40,
            width: 200,
            height: 40,
            style: { fontFamily: "MissingFamily", fontSize: 16 },
          },
          "System font text",
        ),
        createElement(
          Text,
          {
            id: "missing-glyph-fixed",
            x: 520,
            y: 100,
            width: 200,
            height: 40,
            style: { fontFamily: jostRegularOnly, fontSize: 16 },
          },
          "Hello 中文",
        ),
        createElement(
          Text,
          {
            id: "mixed-unmeasured-fixed",
            x: 520,
            y: 160,
            width: 200,
            height: 40,
            style: { fontFamily: jostRegularOnly, fontSize: 16 },
          },
          createElement(Span, { style: { fontFamily: "MissingMix" } }, "ABC"),
          createElement(Span, {}, "中文"),
        ),
        createElement(
          Shape,
          {
            id: "ellipse-with-text",
            x: 520,
            y: 220,
            width: 200,
            height: 100,
            geometry: "ellipse",
            style: { fontFamily: jostRegularOnly, fontSize: 16, padding: 4 },
          },
          "Inside ellipse",
        ),
        createElement(Shape, {
          id: "empty-shape",
          x: 520,
          y: 340,
          width: 100,
          height: 50,
          fill: "#336699",
        }),
        createElement(
          Text,
          {
            id: "ph-title",
            placeholder: "title",
            style: { fontFamily: jostRegularOnly, fontSize: 24 },
          },
          "Placeholder heading",
        ),
        createElement(Placeholder, {
          id: "ph-empty",
          name: "EmptyPh",
          x: 40,
          y: 400,
          width: 120,
          height: 40,
          nativePlaceholder: { type: "body", idx: 1 },
        }),
        createElement(
          Group,
          { id: "grp" },
          createElement(
            Text,
            {
              id: "grp-text",
              x: 40,
              y: 460,
              width: 200,
              height: 40,
              style: { fontFamily: jostRegularOnly, fontSize: 16 },
            },
            "Grouped text",
          ),
        ),
      ),
    );

    const expanded = await expandDocument(doc);
    const resolution = resolve(expanded);
    const slideObs = resolution.report.slides[0]!;
    const byAuthor = (id: string) =>
      slideObs.objects.find((o) => o.authorId === id);

    // Compact 100px "2025": overflows 110px height, fits 150px height
    const compactOverflow = byAuthor("compact-overflow")?.textMetrics;
    assertMeasured(compactOverflow);
    assert.deepEqual(compactOverflow.available, { width: 460, height: 110 });
    assert.equal(compactOverflow.content.lineCount, 1);
    assert.ok(compactOverflow.content.height > 140);
    assert.equal(compactOverflow.fits, false);

    const compactFits = byAuthor("compact-fits")?.textMetrics;
    assertMeasured(compactFits);
    assert.deepEqual(compactFits.available, { width: 460, height: 150 });
    assert.equal(compactFits.content.lineCount, 1);
    assert.equal(compactFits.fits, true);

    // Content-sized auto text with padding: 8
    const autoTextObj = byAuthor("auto-text")!;
    const box: Rect = autoTextObj.box!;
    const autoMetrics = autoTextObj.textMetrics;
    assertMeasured(autoMetrics);
    assert.equal(autoMetrics.available.width, box.width - 16);
    assert.ok(
      Math.abs(autoMetrics.available.height - (box.height - 16)) < 0.02,
    );
    assert.equal(autoMetrics.fits, true);

    // Empty fixed text: 1 empty line
    const emptyMetrics = byAuthor("empty-fixed")?.textMetrics;
    assertMeasured(emptyMetrics);
    assert.equal(emptyMetrics.content.width, 0);
    assert.equal(emptyMetrics.content.lineCount, 1);
    assert.equal(emptyMetrics.fits, true);

    // Missing font resource
    const missingFontMetrics = byAuthor("missing-font-fixed")?.textMetrics;
    assertUnmeasured(missingFontMetrics);
    assert.equal(missingFontMetrics.reason, "missing-font-resource");
    assert.deepEqual(missingFontMetrics.fonts, [
      { family: "MissingFamily", weight: 400, style: "normal" },
    ]);

    // Missing glyphs
    const missingGlyphMetrics = byAuthor("missing-glyph-fixed")?.textMetrics;
    assertUnmeasured(missingGlyphMetrics);
    assert.equal(missingGlyphMetrics.reason, "missing-glyphs");
    assert.deepEqual(missingGlyphMetrics.fonts, [
      {
        family: "Jost*",
        weight: 400,
        style: "normal",
        missingGlyphs: ["中", "文"],
      },
    ]);

    // Mixed unmeasured precedence: missing-font-resource beats missing-glyphs
    const mixedMetrics = byAuthor("mixed-unmeasured-fixed")?.textMetrics;
    assertUnmeasured(mixedMetrics);
    assert.equal(mixedMetrics.reason, "missing-font-resource");
    assert.deepEqual(mixedMetrics.fonts, [
      { family: "MissingMix", weight: 400, style: "normal" },
    ]);

    // Shape with text deducts preset geometry textRect + padding (ellipse textRect is inscribed ~0.707 of box)
    const ellipseMetrics = byAuthor("ellipse-with-text")?.textMetrics;
    assertMeasured(ellipseMetrics);
    assert.ok(ellipseMetrics.available.width < 200 - 8);
    assert.ok(ellipseMetrics.available.height < 100 - 8);

    // Empty shape and unpopulated Placeholder do not have textMetrics
    assert.equal(byAuthor("empty-shape")?.textMetrics, undefined);
    assert.equal(byAuthor("ph-empty")?.textMetrics, undefined);

    // Text-bound placeholder has textMetrics
    const phTitleMetrics = byAuthor("ph-title")?.textMetrics;
    assertMeasured(phTitleMetrics);

    // Group member has textMetrics
    const grpObj = byAuthor("grp")!;
    assert.equal(grpObj.textMetrics, undefined);
    const grpTextMetrics = grpObj.members?.[0]?.textMetrics;
    assertMeasured(grpTextMetrics);

    // Slide effective footer items expose textMetrics
    assert.ok(slideObs.footer !== undefined);
    assertUnmeasured(slideObs.footer.dateTime?.textMetrics);
    assert.equal(
      slideObs.footer.dateTime?.textMetrics.reason,
      "missing-font-resource",
    );
    assertUnmeasured(slideObs.footer.footer?.textMetrics);
    assertUnmeasured(slideObs.footer.slideNumber?.textMetrics);
  });

  it("drives table cell SVG <clipPath> from TextMetrics and preserves pixel-identical PNG output", async () => {
    const doc = createElement(
      Deck,
      { width: 320, height: 180 },
      createElement(
        Slide,
        { id: "overflow-slide", background: "#FFFFFF" },
        createElement(
          Table,
          {
            id: "tbl-overflow",
            x: 20,
            y: 20,
            columns: [120],
            style: { fontFamily: jostRegularOnly, fontSize: 16 },
          },
          createElement(
            Row,
            { id: "r1", style: { height: 28 } },
            createElement(
              Cell,
              { id: "c1", style: { padding: 4, color: "#CC0000" } },
              "Line 1\nLine 2\nLine 3\nLine 4",
            ),
          ),
        ),
      ),
      createElement(
        Slide,
        { id: "control-slide", background: "#FFFFFF" },
        createElement(
          Table,
          {
            id: "tbl-control",
            x: 20,
            y: 20,
            columns: [120],
            style: { fontFamily: jostRegularOnly, fontSize: 16 },
          },
          createElement(
            Row,
            { id: "r1", style: { height: 40 } },
            createElement(
              Cell,
              { id: "c1", style: { padding: 4, color: "#111111" } },
              "Fits",
            ),
          ),
        ),
      ),
    );

    const previewResult = await preview(doc);
    const p0 = previewResult.pages[0]!;
    const p1 = previewResult.pages[1]!;

    assert.ok(p0.svg?.includes("<clipPath"));
    assert.equal(p1.svg?.includes("<clipPath"), false);
    assert.ok(p0.png !== undefined);
    assert.ok(p1.png !== undefined);

    const overflowSha = createHash("sha256").update(p0.png).digest("hex");
    const controlSha = createHash("sha256").update(p1.png).digest("hex");
    assert.equal(overflowSha, EXPECTED_OVERFLOW_PNG_SHA256);
    assert.equal(controlSha, EXPECTED_CONTROL_PNG_SHA256);

    // Table anchor cells expose TextMetrics on table parts and query locators
    const overflowTable = previewResult.report.slides[0]?.objects.find(
      (o) => o.authorId === "tbl-overflow",
    )?.table;
    const overflowCellPart = overflowTable?.parts?.find(
      (part) => part.kind === "cell" && part.row === 0 && part.column === 0,
    );
    assertMeasured(overflowCellPart?.textMetrics);
    assert.equal(overflowCellPart.textMetrics.fits, false);
    assert.equal(overflowCellPart.textMetrics.content.lineCount, 4);

    const controlCellObs = previewResult.query
      .getSlide("control-slide")
      .get("tbl-control")
      .get("r1")
      .get("c1")
      .observe();
    assertMeasured(controlCellObs.textMetrics);
    assert.equal(controlCellObs.textMetrics.fits, true);
  });
});

describe("Ticket 01 Slice 2: document-level FontUsage summary across all surfaces", () => {
  it("aggregates fonts on ResolutionReport and BuildReport across text, shape, placeholder, group, table, chart, notes, footer, and master/layout objects", async () => {
    const restrictedFont = font("JostRestricted", [
      { weight: 400, data: withFontFsType(jostBookBytes, 0x0002) },
    ]);
    const nonEmbeddableFlagFont = font(
      "JostNoEmbed",
      [{ weight: 400, data: jostBookBytes }],
      { embeddable: false },
    );
    const substituteFont = font(
      "SubstitutedSans",
      [{ weight: 400, data: jostBookBytes }],
      { substitute: true },
    );

    const layoutDef = defineSlideLayout({
      id: "summary-layout",
      placeholders: {
        title: {
          type: "title",
          x: 40,
          y: 20,
          width: 400,
          height: 60,
        },
      },
      footerPlaceholders: {
        footer: { style: { fontFamily: "FooterFont", fontSize: 12 } },
      },
    });

    const doc = createElement(
      Deck,
      { width: 960, height: 540 },
      createElement(
        SlideMaster,
        { id: "m1" },
        createElement(
          Text,
          {
            id: "master-badge",
            x: 20,
            y: 10,
            width: 120,
            height: 24,
            style: { fontFamily: "MasterSharedFont", fontSize: 12 },
          },
          "Master",
        ),
        createElement(
          SlideLayout,
          { layout: layoutDef },
          createElement(
            Text,
            {
              id: "layout-badge",
              x: 160,
              y: 10,
              width: 120,
              height: 24,
              style: { fontFamily: "LayoutSharedFont", fontSize: 12 },
            },
            "Layout",
          ),
        ),
      ),
      createElement(
        Slide,
        {
          id: "s1",
          layout: layoutDef,
          footer: "Confidential",
        },
        createElement(
          Text,
          {
            id: "ph-title",
            placeholder: "title",
            style: { fontFamily: jostFont, fontSize: 24 },
          },
          "Title ",
          createElement(Span, { style: { fontWeight: "bold" } }, "Bold"),
        ),
        createElement(
          Text,
          {
            id: "missing-glyph-text",
            x: 40,
            y: 90,
            width: 200,
            height: 40,
            style: { fontFamily: jostRegularOnly, fontSize: 16 },
          },
          "Missing 中",
        ),
        createElement(
          Text,
          {
            id: "restricted-text",
            x: 250,
            y: 90,
            width: 200,
            height: 40,
            style: { fontFamily: restrictedFont, fontSize: 16 },
          },
          "Restricted",
        ),
        createElement(
          Shape,
          {
            id: "shape-text",
            x: 40,
            y: 140,
            width: 200,
            height: 40,
            style: { fontFamily: nonEmbeddableFlagFont, fontSize: 14 },
          },
          "Shape copy",
        ),
        createElement(
          Group,
          { id: "grp" },
          createElement(
            Text,
            {
              id: "grp-text",
              x: 40,
              y: 190,
              width: 200,
              height: 30,
              style: { fontFamily: substituteFont, fontSize: 14 },
            },
            "Group copy",
          ),
        ),
        createElement(
          Table,
          {
            id: "tbl",
            x: 40,
            y: 230,
            columns: [160],
            style: { fontFamily: "TableFont", fontSize: 12 },
          },
          createElement(
            Row,
            { id: "r1", style: { height: 28 } },
            createElement(Cell, { id: "c1" }, "Table cell"),
          ),
        ),
        createElement(Chart, {
          id: "cht",
          x: 300,
          y: 90,
          width: 300,
          height: 200,
          title: {
            text: "Revenue 折",
            style: { fontFamily: jostRegularOnly, fontSize: 16 },
          },
          style: { fontFamily: "ChartBodyFont", fontSize: 12 },
          groups: [
            {
              kind: "column",
              data: { kind: "shared-categories", categories: ["Q1"] },
              series: [{ name: "Sales", values: [10] }],
            },
          ],
        }),
        createElement(
          Notes,
          {},
          createElement(
            P,
            { style: { fontFamily: "NotesFont", fontSize: 12 } },
            "Speaker notes",
          ),
        ),
      ),
    );

    const expanded = await expandDocument(doc);
    const resolution = resolve(expanded);
    const built = await build(expanded);

    const checkFonts = (fonts: readonly FontUsage[]) => {
      const find = (family: string, weight = 400) =>
        fonts.find(
          (f) =>
            f.family.toLowerCase() === family.toLowerCase() &&
            f.weight === weight,
        );

      const masterFont = find("MasterSharedFont");
      assert.ok(masterFont);
      assert.equal(masterFont.hasResource, false);
      assert.equal(masterFont.embeddable, undefined);
      assert.ok(masterFont.paths.some((p) => p.includes("master-badge")));

      const layoutFont = find("LayoutSharedFont");
      assert.ok(layoutFont);
      assert.equal(layoutFont.hasResource, false);
      assert.ok(layoutFont.paths.some((p) => p.includes("layout-badge")));

      const footerFont = find("FooterFont");
      assert.ok(footerFont);
      assert.equal(footerFont.hasResource, false);
      assert.ok(footerFont.paths.length > 0);

      const jost400 = find("Jost*", 400);
      assert.ok(jost400);
      assert.equal(jost400.hasResource, true);
      assert.equal(jost400.embeddable, true);
      assert.deepEqual(jost400.missingGlyphs, ["中", "折"]);
      assert.ok(jost400.paths.some((p) => p.includes("ph-title")));
      assert.ok(jost400.paths.some((p) => p.includes("missing-glyph-text")));
      assert.ok(jost400.paths.some((p) => p.includes("cht")));

      const restricted = find("JostRestricted", 400);
      assert.ok(restricted);
      assert.equal(restricted.hasResource, true);
      assert.equal(restricted.embeddable, false);
      assert.ok(restricted.paths.some((p) => p.includes("restricted-text")));

      const jost700 = find("Jost*", 700);
      assert.ok(jost700);
      assert.equal(jost700.hasResource, true);
      assert.equal(jost700.embeddable, true);
      assert.equal(jost700.missingGlyphs, undefined);

      const noEmbed = find("JostNoEmbed");
      assert.ok(noEmbed);
      assert.equal(noEmbed.hasResource, true);
      assert.equal(noEmbed.embeddable, false);

      const sub = find("SubstitutedSans");
      assert.ok(sub);
      assert.equal(sub.hasResource, true);
      assert.equal(sub.substitute, true);
      assert.equal(sub.embeddable, false);
      assert.ok(sub.paths.some((p) => p.includes("grp-text")));

      const tableFont = find("TableFont");
      assert.ok(tableFont);
      assert.equal(tableFont.hasResource, false);
      assert.ok(tableFont.paths.some((p) => p.includes("tbl")));

      const chartBodyFont = find("ChartBodyFont");
      assert.ok(chartBodyFont);
      assert.equal(chartBodyFont.hasResource, false);
      assert.ok(chartBodyFont.paths.some((p) => p.includes("cht")));

      const notesFont = find("NotesFont");
      assert.ok(notesFont);
      assert.equal(notesFont.hasResource, false);
      assert.ok(notesFont.paths.some((p) => p.includes("slide[s1]")));
    };

    checkFonts(resolution.report.fonts);
    checkFonts(built.report.fonts);
  });
});

describe("Ticket 01 Slice 3: resolve and emit font/glyph hard-error boundaries", () => {
  it("throws BuildError in resolve when content-sized Text, Shape, Placeholder, auto-track Table, or baseline container lacks font bytes or visible glyphs", async () => {
    const assertResolveThrows = async (
      jsx: ReturnType<typeof createElement>,
      expectedCodes: readonly string[],
    ) => {
      const expanded = await expandDocument(jsx);
      assert.throws(
        () => resolve(expanded),
        (error: unknown) =>
          error instanceof BuildError &&
          error.diagnostics.some((d) => expectedCodes.includes(d.code)),
      );
    };

    // Content-sized Text: missing font bytes & missing glyphs
    await assertResolveThrows(
      createElement(
        Deck,
        { width: 480, height: 270 },
        createElement(
          Slide,
          { id: "s1" },
          createElement(
            Text,
            {
              id: "t1",
              x: 20,
              y: 20,
              width: 200,
              style: { fontFamily: "MissingFont", fontSize: 16 },
            },
            "Hello",
          ),
        ),
      ),
      ["font-measurement-unavailable"],
    );
    await assertResolveThrows(
      createElement(
        Deck,
        { width: 480, height: 270 },
        createElement(
          Slide,
          { id: "s1" },
          createElement(
            Text,
            {
              id: "t1",
              x: 20,
              y: 20,
              width: 200,
              style: { fontFamily: jostRegularOnly, fontSize: 16 },
            },
            "中文",
          ),
        ),
      ),
      ["font-missing-glyph"],
    );

    // Content-sized Text Placeholder: missing glyphs
    const autoPhLayout = defineSlideLayout({
      id: "auto-ph-layout",
      placeholders: {
        body: { type: "body", x: 20, y: 20, width: 200, height: 100 },
      },
    });
    await assertResolveThrows(
      createElement(
        Deck,
        { width: 480, height: 270 },
        createElement(
          SlideMaster,
          { id: "m1" },
          createElement(SlideLayout, { layout: autoPhLayout }),
        ),
        createElement(
          Slide,
          { id: "s1", layout: autoPhLayout },
          createElement(
            Text,
            {
              id: "ph1",
              placeholder: "body",
              height: "auto",
              style: { fontFamily: jostRegularOnly, fontSize: 16 },
            },
            "中文",
          ),
        ),
      ),
      ["font-missing-glyph"],
    );

    // Content-sized Shape in flex column: missing font bytes & missing glyphs
    await assertResolveThrows(
      createElement(
        Deck,
        { width: 480, height: 270 },
        createElement(
          Slide,
          { id: "s1" },
          createElement(
            Box,
            { style: { display: "flex", flexDirection: "column", width: 240 } },
            createElement(
              Shape,
              {
                id: "sh1",
                style: { fontFamily: jostRegularOnly, fontSize: 16 },
              },
              "中文",
            ),
          ),
        ),
      ),
      ["font-missing-glyph"],
    );

    // Table with auto column: missing font bytes & missing glyphs
    await assertResolveThrows(
      createElement(
        Deck,
        { width: 480, height: 270 },
        createElement(
          Slide,
          { id: "s1" },
          createElement(
            Table,
            {
              id: "tbl",
              x: 20,
              y: 20,
              columns: ["auto"],
              style: { fontFamily: jostRegularOnly, fontSize: 14 },
            },
            createElement(
              Row,
              { id: "r1", style: { height: 32 } },
              createElement(Cell, { id: "c1" }, "中文"),
            ),
          ),
        ),
      ),
      ["font-missing-glyph"],
    );

    // Table with auto-height Row: missing font bytes & missing glyphs
    await assertResolveThrows(
      createElement(
        Deck,
        { width: 480, height: 270 },
        createElement(
          Slide,
          { id: "s1" },
          createElement(
            Table,
            {
              id: "tbl",
              x: 20,
              y: 20,
              columns: [160],
              style: { fontFamily: jostRegularOnly, fontSize: 14 },
            },
            createElement(
              Row,
              { id: "r1" },
              createElement(Cell, { id: "c1" }, "中文"),
            ),
          ),
        ),
      ),
      ["font-missing-glyph"],
    );

    // Baseline-aligned container: missing font bytes & missing glyphs
    await assertResolveThrows(
      createElement(
        Deck,
        { width: 480, height: 270 },
        createElement(
          Slide,
          { id: "s1" },
          createElement(
            Box,
            {
              style: {
                display: "flex",
                flexDirection: "row",
                alignItems: "baseline",
                width: 400,
                height: 100,
              },
            },
            createElement(
              Text,
              {
                id: "b1",
                style: {
                  width: 150,
                  height: 40,
                  fontFamily: "MissingBaselineFont",
                  fontSize: 16,
                },
              },
              "ABC",
            ),
          ),
        ),
      ),
      ["baseline-measurement-unavailable", "font-measurement-unavailable"],
    );
    await assertResolveThrows(
      createElement(
        Deck,
        { width: 480, height: 270 },
        createElement(
          Slide,
          { id: "s1" },
          createElement(
            Box,
            {
              style: {
                display: "flex",
                flexDirection: "row",
                alignItems: "baseline",
                width: 400,
                height: 100,
              },
            },
            createElement(
              Text,
              {
                id: "b2",
                style: {
                  width: 150,
                  height: 40,
                  fontFamily: jostRegularOnly,
                  fontSize: 16,
                },
              },
              "中文",
            ),
          ),
        ),
      ),
      ["font-missing-glyph"],
    );

    // Whitespace (including U+2028) does NOT trigger .notdef error on content-sized Text
    const wsExpanded = await expandDocument(
      createElement(
        Deck,
        { width: 480, height: 270 },
        createElement(
          Slide,
          { id: "s1" },
          createElement(
            Text,
            {
              id: "ws-ok",
              x: 20,
              y: 20,
              width: 200,
              style: { fontFamily: jostRegularOnly, fontSize: 16 },
            },
            "Hello\u2028World ",
          ),
        ),
      ),
    );
    const wsRes = resolve(wsExpanded);
    assert.equal(
      wsRes.report.slides[0]?.objects[0]?.textMetrics?.status,
      "measured",
    );
  });

  it("allows fixed-region Text/Shape/Table/Chart/Notes/Footer with missing font or glyphs through resolve and non-embedded emit, but rejects emit with embedFonts", async () => {
    const fixedMissingGlyphsDoc = await expandDocument(
      createElement(
        Deck,
        { width: 640, height: 360 },
        createElement(
          Slide,
          { id: "s1", footer: "页脚" },
          createElement(
            Text,
            {
              id: "t-fixed",
              x: 20,
              y: 20,
              width: 160,
              height: 40,
              style: { fontFamily: jostRegularOnly, fontSize: 16 },
            },
            "中文",
          ),
          createElement(
            Shape,
            {
              id: "sh-fixed",
              x: 20,
              y: 70,
              width: 160,
              height: 40,
              style: { fontFamily: jostRegularOnly, fontSize: 16 },
            },
            "中文",
          ),
          createElement(
            Table,
            {
              id: "tbl-fixed",
              x: 20,
              y: 120,
              columns: [160],
              style: { fontFamily: jostRegularOnly, fontSize: 14 },
            },
            createElement(
              Row,
              { id: "r1", style: { height: 30 } },
              createElement(Cell, { id: "c1" }, "中文"),
            ),
          ),
          createElement(Chart, {
            id: "cht-fixed",
            x: 200,
            y: 20,
            width: 240,
            height: 160,
            title: {
              text: "图表",
              style: { fontFamily: jostRegularOnly, fontSize: 14 },
            },
            style: { fontFamily: jostRegularOnly, fontSize: 12 },
            groups: [
              {
                kind: "column",
                data: { kind: "shared-categories", categories: ["一"] },
                series: [{ name: "系列", values: [5] }],
              },
            ],
          }),
          createElement(
            Notes,
            {},
            createElement(
              P,
              { style: { fontFamily: jostRegularOnly, fontSize: 12 } },
              "备注",
            ),
          ),
        ),
      ),
    );

    const resolution = resolve(fixedMissingGlyphsDoc);
    assert.equal(
      resolution.report.slides[0]?.objects.find((o) => o.authorId === "t-fixed")
        ?.textMetrics?.status,
      "unmeasured",
    );
    const plainEmit = emit(resolution);
    assert.ok(plainEmit.pptx.byteLength > 0);

    for (const mode of ["full", "subset"] as const) {
      assert.throws(
        () => emit(resolution, { embedFonts: mode }),
        (error: unknown) =>
          error instanceof BuildError &&
          error.diagnostics.some((d) => d.code === "font-missing-glyph"),
      );
    }

    // Missing bytes and forbidden permissions fail emit with embedFonts
    const missingBytesRes = resolve(
      await expandDocument(
        createElement(
          Deck,
          { width: 480, height: 270 },
          createElement(
            Slide,
            { id: "s1" },
            createElement(
              Text,
              {
                id: "t1",
                x: 20,
                y: 20,
                width: 160,
                height: 40,
                style: { fontFamily: "MissingFamily", fontSize: 16 },
              },
              "Hello",
            ),
          ),
        ),
      ),
    );
    assert.ok(emit(missingBytesRes).pptx.byteLength > 0);
    assert.throws(
      () => emit(missingBytesRes, { embedFonts: "full" }),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics.some(
          (d) => d.code === "font-embedding-resource-missing",
        ),
    );

    const restrictedBytesRes = resolve(
      await expandDocument(
        createElement(
          Deck,
          { width: 480, height: 270 },
          createElement(
            Slide,
            { id: "s1" },
            createElement(
              Text,
              {
                id: "t1",
                x: 20,
                y: 20,
                width: 160,
                height: 40,
                style: {
                  fontFamily: font("Jost*", [
                    {
                      weight: 400,
                      data: withFontFsType(jostBookBytes, 0x0002),
                    },
                  ]),
                  fontSize: 16,
                },
              },
              "Hello",
            ),
          ),
        ),
      ),
    );
    assert.throws(
      () => emit(restrictedBytesRes, { embedFonts: "subset" }),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics.some((d) => d.code === "font-embedding-restricted"),
    );
  });
});
