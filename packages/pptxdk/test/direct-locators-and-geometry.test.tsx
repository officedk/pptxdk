import assert from "node:assert/strict";
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";

import {
  BuildQueryError,
  build,
  defineSlideLayout,
  emit,
  fade,
  font,
  motionSchedule,
  observeMotion,
  resolve,
  type Locator,
  type MotionSchedule,
  type PreviewNode,
  type Rect,
  type ResolvedFooter,
  type ResolvedLayout,
  type ResolvedMaster,
  type ResolvedObject,
  type ResolvedPlaceholder,
  type ResolvedSlide,
  type ResolvedSlideLayout,
  type ResolvedTable,
  type ResolvedTableCell,
  type ResolvedTableRow,
  type SlideLocator,
} from "@officedk/pptxdk";
import { previewFromSource } from "@officedk/pptxdk/node";
import { paint, preview } from "@officedk/pptxdk/preview";
import {
  Box,
  Cell,
  Deck,
  Group,
  Image,
  Row,
  Shape,
  Slide,
  SlideLayout,
  SlideMaster,
  Table,
  Text,
  expandDocument,
} from "@officedk/pptxdk/react";
import { createElement } from "react";
import { describe, it } from "vitest";

const jostBookBytes = new Uint8Array(
  readFileSync(new URL("./assets/fonts/Jost-400-Book.ttf", import.meta.url)),
);

const jostFont = font("Jost*", [{ weight: 400, data: jostBookBytes }]);

const pixelPng = new Uint8Array(
  Buffer.from(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/n1sAAAAASUVORK5CYII=",
    "base64",
  ),
);

describe("Ticket 02: direct locators, Resolved* geometry, and motionSchedule", () => {
  it("exposes slides, slide(id), report.dimensions, and direct Locator properties across resolve, build, preview, and previewFromSource", async () => {
    const jsxDeck = createElement(
      Deck,
      { width: 640, height: 360, title: "Direct Locators" },
      createElement(
        Slide,
        { id: "intro" },
        createElement(
          Box,
          {
            id: "header",
            style: {
              position: "absolute",
              left: 24,
              top: 20,
              width: 300,
              height: 80,
            },
          },
          createElement(
            Text,
            {
              id: "headline",
              x: 0,
              y: 0,
              width: 240,
              height: 40,
              style: { fontFamily: jostFont, fontSize: 18 },
            },
            "Hello direct locator",
          ),
        ),
        createElement(Image, {
          id: "badge",
          x: 400,
          y: 30,
          width: 48,
          height: 48,
          src: pixelPng,
        }),
      ),
    );

    const resolution = resolve(await expandDocument(jsxDeck));
    assert.deepEqual(resolution.report.dimensions, { width: 640, height: 360 });
    assert.equal(resolution.slides.length, 1);

    const resSlide: SlideLocator = resolution.slide("intro");
    assert.equal(resSlide.id, "intro");
    assert.equal(resSlide.index, 1);
    assert.equal(resSlide.path, "deck/slide[intro]");
    assert.deepEqual(resSlide.dimensions, { width: 640, height: 360 });
    assert.equal(resSlide.objects.length, 2);

    const headerLocator: Locator = resSlide.get("header");
    assert.equal(headerLocator.id, "header");
    assert.equal(headerLocator.kind, "layout");
    assert.equal(headerLocator.path, "deck/slide[intro]/box[header]");
    assert.deepEqual(headerLocator.box, {
      x: 24,
      y: 20,
      width: 300,
      height: 80,
    });
    assert.deepEqual(headerLocator.bounds, {
      x: 24,
      y: 20,
      width: 300,
      height: 80,
    });
    assert.equal(headerLocator.text, "Hello direct locator");
    assert.equal(headerLocator.objects.length, 1);
    assert.equal(headerLocator.resources, undefined);

    const headlineLocator = headerLocator.get("headline");
    assert.equal(headlineLocator.id, "headline");
    assert.equal(headlineLocator.kind, "text");
    assert.equal(
      headlineLocator.path,
      "deck/slide[intro]/box[header]/text[headline]",
    );
    const headlineBox: Rect | undefined = headlineLocator.box;
    assert.deepEqual(headlineBox, { x: 24, y: 20, width: 240, height: 40 });
    assert.deepEqual(headlineLocator.bounds, headlineBox);
    assert.equal(headlineLocator.text, "Hello direct locator");
    assert.equal(headlineLocator.textMetrics?.status, "measured");
    assert.equal(headlineLocator.textMetrics?.fits, true);

    // EmitResult stays { pptx, report } and BuildReport carries dimensions.
    const emitted = emit(resolution);
    assert.deepEqual(Object.keys(emitted).toSorted(), ["pptx", "report"]);
    assert.deepEqual(emitted.report.dimensions, { width: 640, height: 360 });

    // PaintResult stays { pages }.
    const painted = paint(resolution);
    assert.deepEqual(Object.keys(painted), ["pages"]);

    // BuildResult populates post-emit resources on Locator directly.
    const built = await build(jsxDeck);
    assert.deepEqual(built.report.dimensions, { width: 640, height: 360 });
    const builtSlide = built.slide("intro");
    const badgeLocator = builtSlide.get("badge");
    assert.equal(badgeLocator.kind, "image");
    assert.ok((badgeLocator.resources?.length ?? 0) >= 1);
    assert.equal(badgeLocator.resources?.[0]?.contentType, "image/png");
    // Transitional bridges remain intact until Ticket 11.
    assert.equal(
      built.query.getSlide("intro").get("badge").observe().authorId,
      "badge",
    );

    // PreviewResult & previewFromSource expose slides & slide(id).
    const previewed = await preview(jsxDeck);
    assert.equal(previewed.slides.length, 1);
    assert.equal(
      previewed.slide("intro").get("header").get("headline").text,
      "Hello direct locator",
    );

    const tempDir = mkdtempSync(join(tmpdir(), "pptxdk-t02-preview-src-"));
    try {
      const entry = join(tempDir, "deck.ts");
      writeFileSync(
        entry,
        [
          `import { deck, slide, text } from ${JSON.stringify(
            new URL("../src/vanilla.ts", import.meta.url).href,
          )};`,
          `export default deck({ width: 480, height: 270 }).append(`,
          `  slide({ id: "cover" }).append(`,
          `    text("From source", { id: "title", x: 10, y: 12, width: 200, height: 30 }),`,
          `  ),`,
          `);`,
        ].join("\n"),
        "utf8",
      );
      const fromSource = await previewFromSource(entry);
      assert.equal(fromSource.slides.length, 1);
      assert.equal(fromSource.slide("cover").get("title").text, "From source");
    } finally {
      rmSync(tempDir, { recursive: true, force: true });
    }

    assert.throws(
      () => resolution.slide("missing"),
      (err: unknown) =>
        err instanceof BuildQueryError && err.code === "missing-slide",
    );
    assert.throws(
      () => resSlide.get("missing"),
      (err: unknown) =>
        err instanceof BuildQueryError && err.code === "missing-match",
    );

    // Type-level smoke check for public Resolved* types.
    const reportSlide: ResolvedSlide = resolution.report.slides[0]!;
    const reportObj: ResolvedObject = reportSlide.objects[0]!;
    const reportLayout: ResolvedLayout | undefined = reportSlide.layouts?.[0];
    const _master: ResolvedMaster | undefined = resolution.report.masters?.[0];
    const _slideLayout: ResolvedSlideLayout | undefined = _master?.layouts[0];
    const _placeholder: ResolvedPlaceholder | undefined =
      _master?.placeholders[0];
    const _footer: ResolvedFooter | undefined = reportSlide.footer;
    assert.ok(reportObj.path.length > 0);
    assert.ok(reportLayout === undefined || reportLayout.path.length > 0);
  });

  it("computes bounds, corners, and transform during resolve for rotated/flipped objects and group members without calling paint", async () => {
    const jsxDeck = createElement(
      Deck,
      { width: 640, height: 360 },
      createElement(
        Slide,
        { id: "geom" },
        createElement(Shape, {
          id: "plain",
          x: 10,
          y: 20,
          width: 80,
          height: 40,
          fill: "#111111",
        }),
        createElement(Shape, {
          id: "rot90",
          x: 100,
          y: 100,
          width: 80,
          height: 40,
          fill: "#222222",
          style: { rotate: 90 },
        }),
        createElement(Shape, {
          id: "flipped",
          x: 240,
          y: 50,
          width: 60,
          height: 30,
          fill: "#333333",
          flipH: true,
        }),
        createElement(
          Group,
          {
            id: "rot-group",
            style: { rotate: 90 },
          },
          createElement(Shape, {
            id: "member",
            x: 300,
            y: 120,
            width: 100,
            height: 40,
            fill: "#444444",
          }),
          createElement(Shape, {
            id: "member2",
            x: 340,
            y: 180,
            width: 60,
            height: 20,
            fill: "#555555",
          }),
        ),
      ),
    );

    // Pure resolve() — no paint() called!
    const resolution = resolve(await expandDocument(jsxDeck));
    const s = resolution.slide("geom");

    const plain = s.get("plain");
    assert.deepEqual(plain.box, { x: 10, y: 20, width: 80, height: 40 });
    assert.deepEqual(plain.bounds, plain.box);
    assert.equal(plain.corners, undefined);
    assert.equal(plain.transform, undefined);

    // 80x40 at (100, 100), center (140, 120), rotated 90 deg -> 40x80 at (120, 80).
    const rot90 = s.get("rot90");
    assert.deepEqual(rot90.box, { x: 100, y: 100, width: 80, height: 40 });
    assert.deepEqual(rot90.bounds, { x: 120, y: 80, width: 40, height: 80 });
    assert.deepEqual(rot90.transform, {
      angle: 90,
      flipH: false,
      flipV: false,
    });
    assert.deepEqual(rot90.corners, [
      { x: 160, y: 80 },
      { x: 160, y: 160 },
      { x: 120, y: 160 },
      { x: 120, y: 80 },
    ]);

    // Flipped horizontally without rotation: bounds equals box, corners mirrored horizontally, transform recorded.
    const flipped = s.get("flipped");
    assert.deepEqual(flipped.bounds, { x: 240, y: 50, width: 60, height: 30 });
    assert.deepEqual(flipped.transform, {
      angle: 0,
      flipH: true,
      flipV: false,
    });
    assert.deepEqual(flipped.corners, [
      { x: 300, y: 50 },
      { x: 240, y: 50 },
      { x: 240, y: 80 },
      { x: 300, y: 80 },
    ]);

    // Group union box: x=300..400 (width 100), y=120..200 (height 80), center=(350, 160).
    // Rotated 90 deg around (350, 160): member (300..400, 120..160) -> (350..390, 110..210).
    const grp = s.get("rot-group");
    assert.deepEqual(grp.bounds, { x: 310, y: 110, width: 80, height: 100 });
    assert.deepEqual(grp.transform, { angle: 90, flipH: false, flipV: false });

    const member = grp.get("member");
    assert.deepEqual(member.box, { x: 300, y: 120, width: 100, height: 40 });
    assert.deepEqual(member.bounds, { x: 350, y: 110, width: 40, height: 100 });
    assert.deepEqual(member.transform, {
      angle: 0,
      flipH: false,
      flipV: false,
    });
    assert.deepEqual(member.corners, [
      { x: 390, y: 110 },
      { x: 390, y: 210 },
      { x: 350, y: 210 },
      { x: 350, y: 110 },
    ]);

    // Also verify PreviewNode carries matching bounds + transitional screenBox.
    const painted = paint(resolution);
    const previewRotNode: PreviewNode | undefined =
      painted.pages[0]?.nodes.find((node) => node.authorId === "rot90");
    assert.deepEqual(previewRotNode?.bounds, rot90.bounds);
    assert.deepEqual(previewRotNode?.screenBox, rot90.bounds);
  });

  it("unconditionally populates ResolvedTable rows and anchor cells with box, span, text, textMetrics, and row/cell Locators", async () => {
    const jsxDeck = createElement(
      Deck,
      { width: 640, height: 360 },
      createElement(
        Slide,
        { id: "tbl-slide" },
        createElement(
          Table,
          {
            x: 20,
            y: 30,
            columns: [100, 120, 80],
            style: { fontFamily: jostFont, fontSize: 14 },
          },
          createElement(
            Row,
            { style: { height: 36 } },
            createElement(Cell, { colSpan: 2 }, "Merged Header"),
            createElement(Cell, null, "Right"),
          ),
          createElement(
            Row,
            { style: { height: 44 } },
            createElement(Cell, null, "A1"),
            createElement(Cell, null, "B1"),
            createElement(Cell, null, "C1"),
          ),
        ),
        createElement(
          Table,
          {
            id: "named-tbl",
            x: 20,
            y: 140,
            columns: [100, 100],
            style: { fontFamily: jostFont, fontSize: 14 },
          },
          createElement(
            Row,
            { id: "r1", style: { height: 30 } },
            createElement(Cell, { id: "c1" }, "In named row"),
            createElement(Cell, null, "Sibling"),
          ),
          createElement(
            Row,
            { style: { height: 30 } },
            createElement(Cell, { id: "c2" }, "In anon row"),
            createElement(Cell, null, "End"),
          ),
        ),
      ),
    );

    const resolution = resolve(await expandDocument(jsxDeck));
    const slideReport = resolution.report.slides[0]!;
    const anonTableObj = slideReport.objects[0]!;
    assert.equal(anonTableObj.kind, "table");
    const tbl: ResolvedTable | undefined = anonTableObj.table;
    assert.ok(tbl !== undefined);

    assert.deepEqual(tbl.columns, [100, 120, 80]);
    assert.deepEqual(tbl.merges, [
      { row: 0, column: 0, rowSpan: 1, colSpan: 2 },
    ]);

    const rows: readonly ResolvedTableRow[] = tbl.rows;
    assert.equal(rows.length, 2);
    assert.deepEqual(rows[0], {
      index: 0,
      path: `${anonTableObj.path}/row[0]`,
      y: 30,
      height: 36,
      box: { x: 20, y: 30, width: 300, height: 36 },
    });
    assert.deepEqual(rows[1], {
      index: 1,
      path: `${anonTableObj.path}/row[1]`,
      y: 66,
      height: 44,
      box: { x: 20, y: 66, width: 300, height: 44 },
    });

    // Anchor cells only (5 anchors for 2x3 grid with 1x2 merge), unconditionally populated without ids!
    const cells: readonly ResolvedTableCell[] = tbl.cells;
    assert.equal(cells.length, 5);
    assert.equal(cells[0]?.row, 0);
    assert.equal(cells[0]?.column, 0);
    assert.equal(cells[0]?.rowSpan, 1);
    assert.equal(cells[0]?.colSpan, 2);
    assert.equal(cells[0]?.path, `${anonTableObj.path}/row[0]/cell[0]`);
    assert.deepEqual(cells[0]?.box, { x: 20, y: 30, width: 220, height: 36 });
    assert.equal(cells[0]?.text, "Merged Header");
    assert.equal(cells[0]?.textMetrics.status, "measured");
    assert.equal(cells[0]?.textMetrics.fits, true);
    assert.equal(cells[1]?.row, 0);
    assert.equal(cells[1]?.column, 2);
    assert.equal(cells[1]?.text, "Right");
    assert.equal(cells[4]?.row, 1);
    assert.equal(cells[4]?.column, 2);
    assert.equal(cells[4]?.text, "C1");

    // Locator traversal over anonymous table -> rows -> anchor cells
    const slideLoc = resolution.slide("tbl-slide");
    const anonTableLoc = slideLoc.children()[0]!;
    assert.equal(anonTableLoc.kind, "table");
    const rowLocs = anonTableLoc.children();
    assert.equal(rowLocs.length, 2);
    assert.equal(rowLocs[0]?.kind, "row");
    assert.equal(rowLocs[0]?.path, `${anonTableObj.path}/row[0]`);
    assert.deepEqual(rowLocs[0]?.box, { x: 20, y: 30, width: 300, height: 36 });
    assert.deepEqual(rowLocs[0]?.bounds, {
      x: 20,
      y: 30,
      width: 300,
      height: 36,
    });
    assert.equal(rowLocs[0]?.text, "Merged Header\t\tRight");

    const row0Cells = rowLocs[0]!.children();
    assert.equal(row0Cells.length, 2);
    assert.equal(row0Cells[0]?.kind, "cell");
    assert.equal(row0Cells[0]?.path, `${anonTableObj.path}/row[0]/cell[0]`);
    assert.equal(row0Cells[0]?.text, "Merged Header");
    assert.equal(row0Cells[0]?.textMetrics?.status, "measured");

    // Scoped id lookup across named vs anonymous rows
    const namedTableLoc = slideLoc.get("named-tbl");
    assert.equal(namedTableLoc.get("r1").get("c1").text, "In named row");
    assert.equal(namedTableLoc.get("c2").text, "In anon row");
  });

  it("exposes resolved columnWidths and rowHeights on grid ResolvedLayout and gridArea on direct grid children", async () => {
    const jsxDeck = createElement(
      Deck,
      { width: 640, height: 360 },
      createElement(
        Slide,
        { id: "grid-slide" },
        createElement(
          Box,
          {
            id: "grid",
            style: {
              position: "absolute",
              left: 40,
              top: 20,
              width: 350,
              height: 150,
              padding: 10,
              display: "grid",
              gridTemplateColumns: "100px 1fr 2fr",
              gridTemplateRows: "40px 1fr",
              columnGap: 10,
              rowGap: 10,
            },
          },
          createElement(Shape, {
            id: "cell-shape",
            style: {
              width: 60,
              height: 20,
              justifySelf: "center",
              alignSelf: "center",
              gridColumn: "1",
              gridRow: "1",
            },
            fill: "#2563EB",
          }),
          createElement(
            Box,
            {
              id: "cell-box",
              style: {
                gridColumn: "2 / 4",
                gridRow: "1",
              },
            },
            createElement(
              Text,
              {
                id: "grandchild",
                style: { fontFamily: jostFont, fontSize: 12 },
              },
              "Inside nested box",
            ),
          ),
          createElement(
            Text,
            {
              id: "cell-text",
              style: {
                gridColumn: "1",
                gridRow: "2",
                fontFamily: jostFont,
                fontSize: 12,
              },
            },
            "Row 2",
          ),
        ),
      ),
    );

    const resolution = resolve(await expandDocument(jsxDeck));
    const slideReport = resolution.report.slides[0]!;
    const gridLayout = slideReport.layouts?.find((l) => l.authorId === "grid");
    assert.ok(gridLayout !== undefined);
    assert.deepEqual(gridLayout.columnWidths, [100, 70, 140]);
    assert.deepEqual(gridLayout.rowHeights, [40, 80]);

    const slideLoc = resolution.slide("grid-slide");
    const shapeLoc = slideLoc.get("grid").get("cell-shape");
    assert.deepEqual(shapeLoc.box, { x: 70, y: 40, width: 60, height: 20 });
    assert.deepEqual(shapeLoc.gridArea, {
      x: 50,
      y: 30,
      width: 100,
      height: 40,
    });
    const shapeObj = slideReport.objects.find(
      (o) => o.authorId === "cell-shape",
    );
    assert.deepEqual(shapeObj?.gridArea, shapeLoc.gridArea);

    const boxLoc = slideLoc.get("grid").get("cell-box");
    assert.deepEqual(boxLoc.gridArea, {
      x: 160,
      y: 30,
      width: 220,
      height: 40,
    });
    const boxLayout = slideReport.layouts?.find(
      (l) => l.authorId === "cell-box",
    );
    assert.deepEqual(boxLayout?.gridArea, boxLoc.gridArea);

    const textLoc = slideLoc.get("grid").get("cell-text");
    assert.deepEqual(textLoc.gridArea, {
      x: 50,
      y: 80,
      width: 100,
      height: 80,
    });

    // Non-grid grandchild inside cell-box has no gridArea.
    assert.equal(boxLoc.get("grandchild").gridArea, undefined);
  });

  it("exposes master/layout shared objects and resolved placeholders, effective slide footer, and motionSchedule(slide)", async () => {
    const coverLayoutDef = defineSlideLayout({
      id: "cover-layout",
      name: "Cover Layout",
      placeholders: {
        title: {
          type: "title",
          y: 30,
          style: { color: "#2563EB" },
        },
        body: {
          type: "body",
          idx: 1,
          sz: "half",
          x: 40,
          y: 100,
          width: 560,
          height: 180,
        },
      },
    });

    const jsxDeck = createElement(
      Deck,
      { width: 640, height: 360 },
      createElement(
        SlideMaster,
        {
          id: "brand-master",
          placeholders: {
            title: {
              type: "title",
              x: 40,
              y: 20,
              width: 560,
              height: 60,
              style: {
                fontFamily: jostFont,
                fontSize: 28,
                color: "#1E293B",
              },
            },
          },
          headerFooter: {
            dateTime: true,
            footer: true,
            slideNumber: true,
          },
          footerPlaceholders: {
            dateTime: {
              x: 40,
              y: 310,
              width: 140,
              height: 24,
              text: "2026-11-25",
              style: { fontFamily: jostFont, fontSize: 10 },
              visible: true,
            },
            footer: {
              x: 200,
              y: 310,
              width: 220,
              height: 24,
              text: "Default Footer",
              style: { fontFamily: jostFont, fontSize: 10 },
              visible: true,
            },
            slideNumber: {
              x: 500,
              y: 310,
              width: 100,
              height: 24,
              style: { fontFamily: jostFont, fontSize: 10 },
              visible: true,
            },
          },
        },
        createElement(Shape, {
          id: "master-bar",
          x: 0,
          y: 0,
          width: 640,
          height: 8,
          fill: "#0F172A",
        }),
        createElement(
          SlideLayout,
          { layout: coverLayoutDef },
          createElement(
            Text,
            {
              id: "layout-tag",
              x: 40,
              y: 286,
              width: 120,
              height: 20,
              style: { fontFamily: jostFont, fontSize: 10 },
            },
            "Layout Tag",
          ),
        ),
      ),
      createElement(
        Slide,
        {
          id: "s1",
          layout: coverLayoutDef,
          footer: "Custom Slide Footer",
          placeholders: {
            title: "Welcome",
            body: "Content",
          },
        },
        createElement(Shape, {
          id: "badge",
          x: 480,
          y: 120,
          width: 80,
          height: 40,
          fill: "#10B981",
          enter: fade({ duration: 400 }),
        }),
      ),
      createElement(Slide, {
        id: "s2",
        layout: coverLayoutDef,
        dateTime: false,
        footer: false,
        slideNumber: false,
      }),
    );

    const resolution = resolve(await expandDocument(jsxDeck));

    // Master shared objects and placeholders
    const master = resolution.report.masters?.[0];
    assert.ok(master !== undefined);
    assert.equal(master.objects.length, 1);
    assert.equal(master.objects[0]?.authorId, "master-bar");
    assert.equal(master.objects[0]?.kind, "shape");
    assert.deepEqual(master.objects[0]?.bounds, {
      x: 0,
      y: 0,
      width: 640,
      height: 8,
    });
    assert.equal(master.placeholders.length, 1);
    assert.equal(master.placeholders[0]?.key, "title");
    assert.equal(master.placeholders[0]?.type, "title");
    assert.deepEqual(master.placeholders[0]?.box, {
      x: 40,
      y: 20,
      width: 560,
      height: 60,
    });
    assert.equal(master.placeholders[0]?.style?.["color"], "#1E293B");

    // Layout shared objects and inherited + overridden placeholders
    const layout = master.layouts[0]!;
    assert.equal(layout.objects.length, 1);
    assert.equal(layout.objects[0]?.authorId, "layout-tag");
    assert.equal(layout.objects[0]?.text, "Layout Tag");
    assert.equal(layout.objects[0]?.textMetrics?.status, "measured");

    const layoutTitle = layout.placeholders.find((ph) => ph.key === "title");
    assert.ok(layoutTitle !== undefined);
    assert.equal(layoutTitle.type, "title");
    assert.deepEqual(layoutTitle.box, {
      x: 40,
      y: 30,
      width: 560,
      height: 60,
    });
    assert.equal(layoutTitle.style?.["color"], "#2563EB");

    const layoutBody = layout.placeholders.find((ph) => ph.key === "body");
    assert.ok(layoutBody !== undefined);
    assert.equal(layoutBody.type, "body");
    assert.equal(layoutBody.idx, 1);
    assert.equal(layoutBody.size, "half");
    assert.deepEqual(layoutBody.box, {
      x: 40,
      y: 100,
      width: 560,
      height: 180,
    });

    // Slide effective footer on report and SlideLocator
    const s1Report = resolution.report.slides[0]!;
    const s1Loc = resolution.slide("s1");
    assert.ok(s1Report.footer !== undefined);
    assert.deepEqual(s1Loc.footer, s1Report.footer);
    assert.equal(s1Loc.footer?.dateTime?.text, "2026-11-25");
    assert.equal(s1Loc.footer?.footer?.text, "Custom Slide Footer");
    assert.equal(s1Loc.footer?.slideNumber?.text, "1");
    assert.equal(s1Loc.footer?.footer?.textMetrics.status, "measured");
    assert.deepEqual(s1Loc.footer?.footer?.box, {
      x: 200,
      y: 310,
      width: 220,
      height: 24,
    });

    const s2Loc = resolution.slide("s2");
    assert.equal(resolution.report.slides[1]?.footer, undefined);
    assert.equal(s2Loc.footer, undefined);

    // motionSchedule accepts both SlideLocator and ResolvedSlide
    const scheduleFromLocator: MotionSchedule | undefined =
      motionSchedule(s1Loc);
    const scheduleFromReport: MotionSchedule | undefined =
      motionSchedule(s1Report);
    assert.ok(scheduleFromLocator !== undefined);
    assert.equal(scheduleFromLocator, scheduleFromReport);
    assert.equal(scheduleFromLocator.steps.length, 1);
    const effectObs = scheduleFromLocator.steps[0]!.effects[0]!;
    assert.equal(effectObs.target.authorId, "badge");
    assert.equal(effectObs.start.state, "known");
    assert.equal(effectObs.start.ms, 0);
    assert.equal(effectObs.end.state, "known");
    assert.equal(effectObs.end.ms, 400);
    assert.equal(observeMotion(s1Loc), scheduleFromLocator);
  });
});
