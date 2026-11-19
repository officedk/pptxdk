import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import { expandDocument } from "@react-pptx/core";
import { createResolver } from "@react-pptx/core/build";
import { paint } from "@react-pptx/core/preview";
import { emit } from "@react-pptx/emit";
import {
  BuildError,
  Cell,
  Deck,
  P,
  Row,
  Slide,
  SlideNumber,
  Table,
  font,
  slideBackgroundCapability,
  tableCapability,
} from "@react-pptx/react-pptx";
import {
  cell,
  deck,
  font as impFont,
  row,
  slide,
  table,
} from "@react-pptx/react-pptx/document";
import { validatePptx } from "@react-pptx/validation";
import { strFromU8, unzipSync } from "fflate";
import { createElement } from "react";
import { describe, it, vi } from "vitest";

const jostBook = readFileSync(
  new URL("./assets/fonts/Jost-400-Book.ttf", import.meta.url),
);
const jostFont = font("Jost*", [{ weight: 400, data: jostBook }]);
const impJostFont = impFont("Jost*", [{ weight: 400, data: jostBook }]);

describe("tableCapability resolved contract", () => {
  it("resolves table layout, cell bodies, and sub-part locators and paints without calling emit", async () => {
    const bindSpy = vi.spyOn(tableCapability, "bind");
    try {
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
      assert.equal(bindSpy.mock.calls.length, 1);
      const boundHandle = bindSpy.mock.results[0]?.value;
      assert.ok(boundHandle !== undefined);
      assert.equal(typeof boundHandle.resolve, "function");

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
      assert.deepEqual(observed.table?.cells, [
        ["Plan Overview", "", "Page 1"],
        ["Pro", "Wraps inside flexible middle track", "$49"],
      ]);

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
    } finally {
      bindSpy.mockRestore();
    }
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

    // Missing glyph in a cell produces cell-located warning during resolve,
    // allows paint and non-embedded emit, and fails with cell path on embedded emit.
    const missingGlyphDoc = deck({ width: 480, height: 270 }).append(
      slide({ id: "cover" }).append(
        table({
          id: "tbl",
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
            code === "invalid-xml-character" &&
            path === "deck/slide[cover]/table[bad-name]",
        ),
    );

    // Cell overflow & unsupported border style fail paint with cell paths while emit still succeeds.
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
    assert.equal(paintedErrors.pages[0]?.png, undefined);
    assert.equal(
      paintedErrors.pages[0]?.diagnostics[0]?.path,
      "deck/slide[overflow-slide]/table[overflow-tbl]/row[r1]/cell[c1]",
    );
    assert.equal(paintedErrors.pages[1]?.png, undefined);
    assert.equal(
      paintedErrors.pages[1]?.diagnostics[0]?.path,
      "deck/slide[border-slide]/table[border-tbl]/row[r1]/cell[c1]",
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
