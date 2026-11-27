import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import {
  cell,
  deck,
  font,
  paragraph,
  row,
  slide,
  span,
  table,
} from "@pptxdk/capabilities/document";
import type { TableColumn } from "@pptxdk/capabilities/table";
import { BuildError } from "@pptxdk/core";
import { createResolver } from "@pptxdk/core/build";
import { CommonContainer, RichTextContainer } from "@pptxdk/core/document";
import { emit } from "@pptxdk/emit";
import { validatePptx } from "@pptxdk/validation";
import { strFromU8, unzipSync } from "fflate";
import { describe, it } from "vitest";

import { build } from "../../../test/test-helpers.ts";

const notoSansScRegular = new Uint8Array(
  readFileSync(
    new URL(
      "../../../test/assets/fonts/NotoSansSC-Regular.woff2",
      import.meta.url,
    ),
  ),
);

const metricFont = font(
  "Table Metric",
  [{ weight: 400, data: notoSansScRegular }],
  { substitute: true },
);

function slideXml(result: { readonly pptx: Uint8Array }, index = 1): string {
  return strFromU8(unzipSync(result.pptx)[`ppt/slides/slide${index}.xml`]!);
}

describe("table authoring through the common document", () => {
  it("creates an editable native table, preserves rich text, spans, styles, and measured tracks, then edits it", async () => {
    const columns: TableColumn[] = ["auto", "1fr"];
    const mergedStyle: { background: `#${string}` } = { background: "#DDEEFF" };
    const merged = cell(
      paragraph().append(
        span("North").style({ fontWeight: "bold", color: "#AA0000" }),
        " Q1",
      ),
      { colSpan: 2, style: mergedStyle },
    );
    const revenue = cell("Revenue");
    const value = cell("42");
    const firstRow = row({ style: { height: 30 } }).append(merged);
    const secondRow = row().append(revenue, value);
    assert.ok(merged instanceof RichTextContainer);
    assert.ok(firstRow instanceof CommonContainer);
    const tableNode = table({
      id: "document-table",
      columns,
      style: {
        position: "absolute",
        left: 16,
        top: 20,
        width: 360,
        fontFamily: metricFont,
        fontSize: 12,
      },
    }).append(firstRow, secondRow);
    assert.ok(tableNode instanceof CommonContainer);
    const document = deck({ width: 960, height: 540 }).append(
      slide({ id: "table" }).append(tableNode),
    );

    // Values are received at construction time; later edits to the caller's inputs do not leak in.
    columns.push("1fr");
    mergedStyle.background = "#0000FF";

    const first = await build(document);
    const firstFile = await validatePptx(first.pptx);
    assert.deepEqual(firstFile.errors, []);
    const tableObject = firstFile.slides[0]?.objects[0];
    assert.equal(tableObject?.kind, "table");
    const tableObservation = tableObject?.table;
    assert.ok(tableObservation !== undefined);
    assert.deepEqual(tableObservation.cells, [
      ["North Q1", ""],
      ["Revenue", "42"],
    ]);
    assert.deepEqual(tableObservation.merges, [
      { row: 0, column: 0, rowSpan: 1, colSpan: 2 },
    ]);
    assert.equal(tableObservation.columnWidthsEmu.length, 2);
    const tableWidth = tableObservation.columnWidthsEmu.reduce(
      (sum, width) => sum + width,
      0,
    );
    assert.ok(Math.abs(tableWidth - 360 * 9525) <= 1);
    assert.equal(tableObservation.rowHeightsEmu[0], 30 * 9525);
    assert.ok((tableObservation.rowHeightsEmu[1] ?? 0) > 0);

    const mergedText = tableObservation.cellTexts.find(
      ({ row: cellRow, column }) => cellRow === 0 && column === 0,
    )?.text;
    assert.equal(mergedText?.text, "North Q1");
    assert.equal(mergedText?.paragraphs[0]?.runs[0]?.text, "North");
    assert.equal(mergedText?.paragraphs[0]?.runs[0]?.bold, true);
    assert.equal(mergedText?.paragraphs[0]?.runs[0]?.color, "#AA0000");
    assert.match(slideXml(first), /DDEEFF/u);
    assert.doesNotMatch(slideXml(first), /0000FF/u);

    merged.content("Updated ", span("cell").style({ fontStyle: "italic" }));
    value.style({ background: "#FFFF00" });
    secondRow.style({ height: 40 });

    const second = await build(document);
    const secondFile = await validatePptx(second.pptx);
    assert.deepEqual(secondFile.errors, []);
    const updatedTable = secondFile.slides[0]?.objects[0]?.table;
    assert.deepEqual(updatedTable?.cells, [
      ["Updated cell", ""],
      ["Revenue", "42"],
    ]);
    assert.equal(updatedTable?.rowHeightsEmu[1], 40 * 9525);
    assert.match(slideXml(second), /Updated /u);
    assert.match(slideXml(second), />cell</u);
    assert.match(slideXml(second), /FFFF00/u);
    assert.match(slideXml(second), /i="1"/u);
    assert.doesNotMatch(slideXml(first), /Updated /u);
    assert.deepEqual(firstFile, await validatePptx(first.pptx));
  }, 10_000);

  it("keeps table registration explicit for trimmed builds", async () => {
    const document = deck({ width: 960, height: 540 }).append(
      slide({ id: "trimmed" }).append(
        table({ columns: [100] }).append(
          row({ style: { height: 20 } }).append(cell("Trimmed table")),
        ),
      ),
    );

    await assert.rejects(
      async () => emit(createResolver([])(document)),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.ok(
          error.diagnostics.some(({ code }) => code === "missing-capability"),
        );
        return true;
      },
    );
  });
});
