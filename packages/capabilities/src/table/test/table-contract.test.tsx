import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import { slideBackgroundCapability } from "@pptxdk/capabilities";
import {
  cell,
  deck,
  font as impFont,
  row,
  slide,
  table,
} from "@pptxdk/capabilities/document";
import { font } from "@pptxdk/capabilities/font";
import { tableCapability } from "@pptxdk/capabilities/table";
import { Cell, Row, Table } from "@pptxdk/capabilities/table-elements";
import { P, SlideNumber } from "@pptxdk/capabilities/text-elements";
import { BuildError } from "@pptxdk/core";
import { createResolver } from "@pptxdk/core/build";
import { emit } from "@pptxdk/emit";
import { paint } from "@pptxdk/preview";
import { Deck, expandDocument, Slide } from "@pptxdk/react";
import { validatePptx } from "@pptxdk/validation";
import { strFromU8, unzipSync } from "fflate";
import { createElement } from "react";
import { describe, it } from "vitest";

const jostBook = readFileSync(
  new URL("../../../test/assets/fonts/Jost-400-Book.ttf", import.meta.url),
);
const jostFont = font("Jost*", [{ weight: 400, data: jostBook }]);
const impJostFont = impFont("Jost*", [{ weight: 400, data: jostBook }]);

describe("tableCapability resolved contract", () => {
  it("resolves table layout, cell bodies, and sub-part locators and paints without calling emit", async () => {
    const resolve = createResolver([
      slideBackgroundCapability,
      tableCapability,
    ]);
    const doc = await expandDocument(
      createElement(
        Deck,
        { width: 480, height: 270 },
        createElement(
          Slide,
          { id: "cover" },
          createElement(
            Table,
            {
              id: "pricing",
              name: "PricingTable",
              x: 20,
              y: 30,
              width: 300,
              columns: [100, "1fr", "auto"],
              style: {
                fontFamily: jostFont,
                fontSize: 14,
              },
            },
            createElement(
              Row,
              {
                id: "header",
                style: { height: 40, backgroundColor: "#1E293B" },
              },
              createElement(
                Cell,
                {
                  id: "title-span",
                  colSpan: 2,
                  style: { color: "#FFFFFF", padding: 6 },
                },
                "Plan Overview",
              ),
              createElement(
                Cell,
                {
                  id: "page-cell",
                  style: { color: "#FFFFFF", padding: 6 },
                },
                createElement(
                  P,
                  null,
                  "Page ",
                  createElement(SlideNumber, null),
                ),
              ),
            ),
            createElement(
              Row,
              { id: "body" },
              createElement(
                Cell,
                {
                  id: "tier-cell",
                  style: {
                    padding: 6,
                    borderColor: "#2563EB",
                    borderWidth: 2,
                  },
                },
                "Pro",
              ),
              createElement(
                Cell,
                { id: "desc-cell", style: { padding: 6 } },
                "Wraps inside flexible middle track",
              ),
              createElement(
                Cell,
                { id: "price-cell", style: { padding: 6 } },
                "$49",
              ),
            ),
          ),
        ),
      ),
    );

    const resolution = resolve(doc);

    const observed = resolution.report.slides[0]?.objects.find(
      (obj) => obj.authorId === "pricing",
    );
    assert.ok(observed !== undefined);
    assert.equal(observed.kind, "table");
    assert.equal(observed.name, "PricingTable");
    assert.equal(observed.table?.columnWidths.length, 3);
    assert.deepEqual(observed.table?.merges, [
      { row: 0, column: 0, rowSpan: 1, colSpan: 2 },
    ]);
    assert.deepEqual(
      observed.table?.cells.map((c) => ({
        row: c.row,
        column: c.column,
        text: c.text,
      })),
      [
        { row: 0, column: 0, text: "Plan Overview" },
        { row: 0, column: 2, text: "Page 1" },
        { row: 1, column: 0, text: "Pro" },
        { row: 1, column: 1, text: "Wraps inside flexible middle track" },
        { row: 1, column: 2, text: "$49" },
      ],
    );

    const tableQuery = resolution.query.getSlide("cover").get("pricing");
    const titleSpanBox = tableQuery
      .get("header")
      .get("title-span")
      .observe().box!;
    assert.equal(
      titleSpanBox.width,
      observed.table.columnWidths[0]! + observed.table.columnWidths[1]!,
    );
    assert.equal(titleSpanBox.height, 40);

    const painted = paint(resolution);
    assert.equal(painted.pages.length, 1);
    assert.deepEqual(painted.pages[0]?.diagnostics, []);
    assert.ok(painted.pages[0]?.png !== undefined);
    assert.ok(painted.pages[0]?.svg?.includes("<path"));
    assert.ok(painted.pages[0]?.svg?.includes("#1E293B"));
    assert.ok(painted.pages[0]?.svg?.includes("#2563EB"));

    const emitted = emit(resolution);
    const validation = await validatePptx(emitted.pptx);
    assert.deepEqual(validation.errors, []);
    const slideXml = strFromU8(
      unzipSync(emitted.pptx)["ppt/slides/slide1.xml"]!,
    );
    assert.ok(slideXml.includes("<a:tbl>"));
    assert.ok(slideXml.includes('type="slidenum"'));
  });

  it("defers font embedding and preserves cell-level diagnostic paths across resolve, emit, and paint", async () => {
    const resolve = createResolver([
      slideBackgroundCapability,
      tableCapability,
    ]);

    const subsetDoc = await expandDocument(
      createElement(
        Deck,
        { width: 480, height: 270 },
        createElement(
          Slide,
          { id: "cover" },
          createElement(
            Table,
            {
              id: "embed-tbl",
              x: 20,
              y: 20,
              columns: [120, 120],
              style: { fontFamily: jostFont, fontSize: 14 },
            },
            createElement(
              Row,
              { id: "r1" },
              createElement(Cell, { id: "c1", style: { padding: 4 } }, "Alpha"),
              createElement(Cell, { id: "c2", style: { padding: 4 } }, "Beta"),
            ),
          ),
        ),
      ),
    );
    const subsetResolution = resolve(subsetDoc);
    const emittedSubset = emit(subsetResolution, { embedFonts: "subset" });
    assert.deepEqual((await validatePptx(emittedSubset.pptx)).errors, []);
    assert.ok(
      Object.keys(unzipSync(emittedSubset.pptx)).some(
        (entry) => entry.startsWith("ppt/fonts/") && entry.endsWith(".fntdata"),
      ),
    );
    const emittedPlain = emit(subsetResolution);
    assert.equal(
      Object.keys(unzipSync(emittedPlain.pptx)).some((entry) =>
        entry.startsWith("ppt/fonts/"),
      ),
      false,
    );

    // Missing glyph in a fixed-height cell produces cell-located warning during resolve,
    // allows paint and non-embedded emit, and fails with cell path on embedded emit;
    // auto-height row with missing glyphs fails immediately in resolve.
    const autoMissingGlyphDoc = deck({ width: 480, height: 270 }).append(
      slide({ id: "cover" }).append(
        table({
          id: "tbl-auto",
          x: 20,
          y: 20,
          columns: [120],
          style: { fontFamily: impJostFont, fontSize: 14 },
        }).append(
          row({ id: "r1" }).append(
            cell("中文", { id: "c1", style: { padding: 4 } }),
          ),
        ),
      ),
    );
    assert.throws(
      () => resolve(autoMissingGlyphDoc),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics.some(
          ({ code, path }) =>
            code === "font-missing-glyph" &&
            path === "deck/slide[cover]/table[tbl-auto]/row[r1]/cell[c1]",
        ),
    );

    const missingGlyphDoc = deck({ width: 480, height: 270 }).append(
      slide({ id: "cover" }).append(
        table({
          id: "tbl",
          x: 20,
          y: 20,
          columns: [120],
          style: { fontFamily: impJostFont, fontSize: 14 },
        }).append(
          row({ id: "r1", style: { height: 32 } }).append(
            cell("中文", { id: "c1", style: { padding: 4 } }),
          ),
        ),
      ),
    );
    const missingResolution = resolve(missingGlyphDoc);
    assert.ok(
      missingResolution.report.diagnostics.some(
        (diag) =>
          diag.severity === "warning" &&
          diag.code === "font-missing-glyph" &&
          diag.path === "deck/slide[cover]/table[tbl]/row[r1]/cell[c1]",
      ),
    );
    assert.deepEqual(
      (await validatePptx(emit(missingResolution).pptx)).errors,
      [],
    );
    assert.throws(
      () => emit(missingResolution, { embedFonts: "full" }),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics.some(
          ({ code, path }) =>
            code === "font-missing-glyph" &&
            path === "deck/slide[cover]/table[tbl]/row[r1]/cell[c1]",
        ),
    );

    // Invalid XML 1.0 characters in Table.name fail in resolve before emit.
    assert.throws(
      () =>
        resolve(
          deck({ width: 480, height: 270 }).append(
            slide({ id: "cover" }).append(
              table({
                id: "bad-name",
                name: "Bad\u0000Table",
                columns: [80],
              }).append(
                row({ id: "r1", style: { height: 30 } }).append(
                  cell("OK", { id: "c1" }),
                ),
              ),
            ),
          ),
        ),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics.some(
          ({ code, path }) =>
            code === "invalid-prop" &&
            path === "deck/slide[cover]/table[bad-name]",
        ),
    );

    // Fixed-height cell overflow produces build overset warning while painting clipped output cleanly; dashed border renders cleanly; missing glyph fails paint with cell path.
    const paintErrorsDoc = deck({ width: 480, height: 270 }).append(
      slide({ id: "overflow-slide" }).append(
        table({
          id: "overflow-tbl",
          x: 20,
          y: 20,
          columns: [80],
          style: { fontFamily: impJostFont, fontSize: 14 },
        }).append(
          row({ id: "r1", style: { height: 18 } }).append(
            cell("Line 1\nLine 2\nLine 3", { id: "c1", style: { padding: 4 } }),
          ),
        ),
      ),
      slide({ id: "border-slide" }).append(
        table({
          id: "border-tbl",
          x: 20,
          y: 20,
          columns: [120],
          style: { fontFamily: impJostFont, fontSize: 14 },
        }).append(
          row({ id: "r1" }).append(
            cell("Dashed", {
              id: "c1",
              style: {
                borderColor: "#1E293B",
                borderWidth: 2,
                borderStyle: "dashed",
              },
            }),
          ),
        ),
      ),
      slide({ id: "glyph-slide" }).append(
        table({
          id: "glyph-tbl",
          x: 20,
          y: 20,
          columns: [120],
          style: { fontFamily: impJostFont, fontSize: 14 },
        }).append(
          row({ id: "r1", style: { height: 28 } }).append(
            cell("汉", { id: "c1" }),
          ),
        ),
      ),
    );
    const paintErrorsResolution = resolve(paintErrorsDoc);
    assert.ok(
      paintErrorsResolution.report.diagnostics.some(
        (diag) =>
          diag.severity === "warning" &&
          diag.code === "overset" &&
          diag.path ===
            "deck/slide[overflow-slide]/table[overflow-tbl]/row[r1]/cell[c1]",
      ),
    );
    const paintedErrors = paint(paintErrorsResolution);
    assert.deepEqual(paintedErrors.pages[0]?.diagnostics, []);
    assert.ok(paintedErrors.pages[0]?.png !== undefined);
    assert.ok(paintedErrors.pages[0]?.svg?.includes("<clipPath"));
    assert.deepEqual(paintedErrors.pages[1]?.diagnostics, []);
    assert.ok(paintedErrors.pages[1]?.png !== undefined);
    assert.ok(paintedErrors.pages[1]?.svg?.includes('stroke-dasharray="8 6"'));
    assert.equal(paintedErrors.pages[2]?.png, undefined);
    assert.equal(
      paintedErrors.pages[2]?.diagnostics[0]?.path,
      "deck/slide[glyph-slide]/table[glyph-tbl]/row[r1]/cell[c1]",
    );
    assert.deepEqual(
      (await validatePptx(emit(paintErrorsResolution).pptx)).errors,
      [],
    );
  });

  it("isolates resolve closures from subsequent editable document mutations", async () => {
    const resolve = createResolver([
      slideBackgroundCapability,
      tableCapability,
    ]);
    const mutableCell = cell("Snapshot cell", {
      id: "mutable-cell",
      style: { padding: 6 },
    });
    const doc = deck({ width: 480, height: 270 }).append(
      slide({ id: "cover" }).append(
        table({
          id: "mutable-table",
          x: 20,
          y: 20,
          columns: [140],
          style: { fontFamily: impJostFont, fontSize: 14 },
        }).append(row({ id: "r1" }).append(mutableCell)),
      ),
    );

    const resolution = resolve(doc);
    mutableCell.content("Mutated cell");

    assert.equal(
      resolution.query
        .getSlide("cover")
        .get("mutable-table")
        .get("r1")
        .get("mutable-cell")
        .observe().text,
      "Snapshot cell",
    );
    const painted = paint(resolution);
    assert.ok(painted.pages[0]?.png !== undefined);
    const emitted = emit(resolution);
    const validation = await validatePptx(emitted.pptx);
    assert.deepEqual(validation.errors, []);
    assert.equal(
      validation.slides[0]?.objects[0]?.table?.cells[0]?.[0],
      "Snapshot cell",
    );
    const slideXml = strFromU8(
      unzipSync(emitted.pptx)["ppt/slides/slide1.xml"]!,
    );
    assert.ok(slideXml.includes("Snapshot cell"));
    assert.equal(slideXml.includes("Mutated cell"), false);
  });
});
