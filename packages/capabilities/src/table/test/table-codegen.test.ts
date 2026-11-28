import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import { documentCapabilities } from "@pptxdk/capabilities/all";
import {
  cell,
  deck,
  defineSlideLayout,
  fade,
  paragraph,
  row,
  slide,
  slideLayout,
  slideMaster,
  slideNumber,
  span,
  table,
  text,
} from "@pptxdk/capabilities/document";
import { Row, Table } from "@pptxdk/capabilities/table-elements";
import { CodegenError, createCodegen } from "@pptxdk/codegen";
import { CommonNode } from "@pptxdk/core/document";
import { createImporter } from "@pptxdk/import";
import { Deck, Slide, expandDocument } from "@pptxdk/react";
import { createElement } from "react";
import { describe, it } from "vitest";

describe("table capability codegen", () => {
  it("generates vanilla and react source for table, row, cell, track modes, merges, borders, rich text, slide numbers, and placeholder bindings", () => {
    const tableLayout = defineSlideLayout({
      id: "table-layout",
      name: "Table Layout",
      placeholders: {
        title: {
          type: "title",
          style: {
            position: "absolute",
            left: 32,
            top: 24,
            width: 576,
            height: 40,
          },
        },
        data: {
          type: "tbl",
          explicitType: true,
          idx: 1,
          style: {
            position: "absolute",
            left: 32,
            top: 80,
            width: 576,
            height: 220,
          },
        },
      },
    });

    const doc = deck({ width: 640, height: 360 }).append(
      slideMaster({ id: "m1", name: "Main" }).append(slideLayout(tableLayout)),
      slide({ id: "s1", layout: tableLayout }).append(
        table({
          id: "pricing",
          name: "PricingTable",
          columns: [130, "auto", "1fr"],
          link: { url: "https://example.com/pricing" },
          style: {
            position: "absolute",
            left: 32,
            top: 80,
            width: 576,
            fontFamily: "Aptos",
            fontSize: 14,
            color: "#0F172A",
            align: "left",
            border: "none",
            borderTop: { color: "#0F172A", width: 2 },
            borderBottom: { color: "#0F172A", width: 2 },
          },
        }).append(
          row({
            id: "hdr",
            style: {
              height: 44,
              background: "#1E293B",
              color: "#FFFFFF",
              align: "center",
              borderBottom: { color: "#64748B", width: 1 },
            },
          }).append(
            cell(paragraph("Overview", { style: { align: "center" } }), {
              id: "hdr-span",
              colSpan: 2,
              style: {
                fontWeight: "bold",
                padding: [6, 10],
                border: { color: "#0F172A", width: 2, style: "solid" },
              },
            }),
            cell(
              paragraph(["Page ", slideNumber()], {
                style: { align: "right" },
              }),
              {
                id: "hdr-page",
                style: {
                  background: "#334155",
                  border: { color: "#0F172A", width: 1, style: "dashed" },
                },
              },
            ),
          ),
          row({ id: "body-1" }).append(
            cell("Tier\nA", {
              id: "tier-a",
              rowSpan: 2,
              style: {
                verticalAlign: "middle",
                background: "#EFF6FF",
                border: { color: "#2563EB", width: 2, style: "dotted" },
              },
            }),
            cell("Metric", { id: "metric-1" }),
            cell(
              [
                paragraph([
                  span("$49", {
                    style: { fontWeight: "bold", color: "#047857" },
                  }),
                  " / mo",
                ]),
              ],
              { id: "price-1", style: { padding: [4, 8, 6, 10] } },
            ),
          ),
          row({ id: "body-2" }).append(
            cell("Summary", { id: "sum-lbl" }),
            cell([], {
              id: "empty-styled",
              style: { background: "#FEF3C7" },
            }),
          ),
        ),
      ),
      slide({ id: "s2", layout: tableLayout }).append(
        table({
          id: "ph-tbl",
          placeholder: "data",
          nativePlaceholder: { type: "tbl", explicitType: true, idx: 1 },
          columns: ["1fr", "2fr"],
        }).append(row().append(cell("A"), cell("B"))),
      ),
    );

    const generate = createCodegen(documentCapabilities);

    const vanilla = generate(doc, { target: "vanilla" });
    const vanillaSource = vanilla.files[vanilla.entry] as string;
    assert.match(
      vanillaSource,
      /import \{[\s\S]*\bcell\b[\s\S]*\brow\b[\s\S]*\btable\b[\s\S]*\} from "@officedk\/pptxdk\/vanilla";/u,
    );
    assert.match(vanillaSource, /columns: \[130, "auto", "1fr"\]/u);
    assert.match(vanillaSource, /columns: \["1fr", "2fr"\]/u);
    assert.match(vanillaSource, /placeholder: "data"/u);
    assert.doesNotMatch(vanillaSource, /nativePlaceholder/u);
    assert.match(vanillaSource, /colSpan: 2/u);
    assert.match(vanillaSource, /rowSpan: 2/u);
    assert.match(vanillaSource, /border:\s*\{[\s\S]*?style:\s*"dashed"/u);
    assert.match(vanillaSource, /border:\s*\{[\s\S]*?style:\s*"dotted"/u);
    assert.match(vanillaSource, /align:\s*"left"/u);
    assert.match(vanillaSource, /borderTop:\s*\{/u);
    assert.match(vanillaSource, /borderBottom:\s*\{/u);
    assert.match(vanillaSource, /cell\(\s*"Tier\\nA",\s*\{/u);
    assert.match(vanillaSource, /cell\(\[\], \{/u);
    assert.match(vanillaSource, /slideNumber\(\)/u);

    const react = generate(doc, { target: "react" });
    const reactSource = react.files[react.entry] as string;
    assert.match(
      reactSource,
      /import \{[\s\S]*\bCell\b[\s\S]*\bRow\b[\s\S]*\bTable\b[\s\S]*\} from "@officedk\/pptxdk\/react";/u,
    );
    assert.match(
      reactSource,
      /<Table[\s\S]*?columns=\{\[130, "auto", "1fr"\]\}/u,
    );
    assert.match(reactSource, /<Table[\s\S]*?placeholder="data"/u);
    assert.match(reactSource, /<Table[\s\S]*?align:\s*"left"/u);
    assert.match(reactSource, /<Table[\s\S]*?borderTop:\s*\{/u);
    assert.match(reactSource, /<Row[\s\S]*?id="hdr"/u);
    assert.match(
      reactSource,
      /<Row[\s\S]*?align:\s*"center"[\s\S]*?borderBottom:\s*\{/u,
    );
    assert.match(reactSource, /<Cell[\s\S]*?colSpan=\{2\}/u);
    assert.match(
      reactSource,
      /<Cell[\s\S]*?rowSpan=\{2\}[\s\S]*?>\s*\{"Tier\\nA"\}\s*<\/Cell>/u,
    );
    assert.match(
      reactSource,
      /<Cell id="empty-styled" style=\{\{ background: "#FEF3C7" \}\} \/>/u,
    );
    assert.match(reactSource, /<SlideNumber \/>/u);
  });

  it("rejects invalid children, unsupported props, and malformed columns with CodegenError and exact paths", async () => {
    const generate = createCodegen(documentCapabilities);

    const badTableChildDeck = deck().append(
      slide({ id: "s1" }).append(
        table({ id: "t1", columns: [100] }).append(
          paragraph("bad", { id: "p1" }),
        ),
      ),
    );
    assert.throws(
      () => generate(badTableChildDeck, { target: "vanilla" }),
      (error: unknown) =>
        error instanceof CodegenError &&
        error.diagnostics[0]?.code === "invalid-document" &&
        error.diagnostics[0]?.path === "deck/slide[s1]/table[t1]/paragraph[p1]",
    );

    const badTableTextDeck = await expandDocument(
      createElement(
        Deck,
        null,
        createElement(
          Slide,
          { id: "s1" },
          createElement(Table, { id: "t1", columns: [100] }, "stray text"),
        ),
      ),
    );
    assert.throws(
      () => generate(badTableTextDeck, { target: "react" }),
      (error: unknown) =>
        error instanceof CodegenError &&
        error.diagnostics[0]?.code === "invalid-document" &&
        error.diagnostics[0]?.path === "deck/slide[s1]/table[t1]",
    );

    const badRowChildDeck = deck().append(
      slide({ id: "s1" }).append(
        table({ id: "t1", columns: [100] }).append(
          row({ id: "r1" }).append(text("bad", { id: "tx1" })),
        ),
      ),
    );
    assert.throws(
      () => generate(badRowChildDeck, { target: "vanilla" }),
      (error: unknown) =>
        error instanceof CodegenError &&
        error.diagnostics[0]?.code === "invalid-document" &&
        error.diagnostics[0]?.path ===
          "deck/slide[s1]/table[t1]/row[r1]/text[tx1]",
    );

    const badCellChildDeck = deck().append(
      slide({ id: "s1" }).append(
        table({ id: "t1", columns: [100] }).append(
          row({ id: "r1" }).append(
            cell({ id: "c1" } as unknown as string).append(
              row({ id: "nested-row" }),
            ),
          ),
        ),
      ),
    );
    assert.throws(
      () => generate(badCellChildDeck, { target: "react" }),
      (error: unknown) =>
        error instanceof CodegenError &&
        error.diagnostics[0]?.code === "invalid-document" &&
        error.diagnostics[0]?.path ===
          "deck/slide[s1]/table[t1]/row[r1]/cell[c1]/row[nested-row]",
    );

    const badTablePropDeck = deck().append(
      slide({ id: "s1" }).append(
        table({
          id: "t1",
          columns: [100],
          unknownProp: true,
        } as unknown as Parameters<typeof table>[0]).append(
          row().append(cell("ok")),
        ),
      ),
    );
    assert.throws(
      () => generate(badTablePropDeck, { target: "vanilla" }),
      (error: unknown) =>
        error instanceof CodegenError &&
        error.diagnostics[0]?.code === "invalid-document" &&
        error.diagnostics[0]?.path === "deck/slide[s1]/table[t1]@unknownProp",
    );

    const badRowPropDeck = deck().append(
      slide({ id: "s1" }).append(
        table({ id: "t1", columns: [100] }).append(
          row({ id: "r1", colSpan: 2 } as unknown as Parameters<
            typeof row
          >[0]).append(cell("ok")),
        ),
      ),
    );
    assert.throws(
      () => generate(badRowPropDeck, { target: "vanilla" }),
      (error: unknown) =>
        error instanceof CodegenError &&
        error.diagnostics[0]?.code === "invalid-document" &&
        error.diagnostics[0]?.path ===
          "deck/slide[s1]/table[t1]/row[r1]@colSpan",
    );

    const badCellPropDeck = deck().append(
      slide({ id: "s1" }).append(
        table({ id: "t1", columns: [100] }).append(
          row({ id: "r1" }).append(
            cell("ok", {
              id: "c1",
              width: 120,
            } as unknown as Parameters<typeof cell>[1]),
          ),
        ),
      ),
    );
    assert.throws(
      () => generate(badCellPropDeck, { target: "react" }),
      (error: unknown) =>
        error instanceof CodegenError &&
        error.diagnostics[0]?.code === "invalid-document" &&
        error.diagnostics[0]?.path ===
          "deck/slide[s1]/table[t1]/row[r1]/cell[c1]@width",
    );

    const emptyColumnsDeck = deck().append(
      slide({ id: "s1" }).append(
        table({
          id: "t1",
          columns: [],
        }).append(row().append(cell("ok"))),
      ),
    );
    assert.throws(
      () => generate(emptyColumnsDeck, { target: "vanilla" }),
      (error: unknown) =>
        error instanceof CodegenError &&
        error.diagnostics[0]?.code === "invalid-document" &&
        error.diagnostics[0]?.path === "deck/slide[s1]/table[t1]@columns",
    );

    const badTrackDeck = deck().append(
      slide({ id: "s1" }).append(
        table({
          id: "t1",
          columns: ["50%"],
        } as unknown as Parameters<typeof table>[0]).append(
          row().append(cell("ok")),
        ),
      ),
    );
    assert.throws(
      () => generate(badTrackDeck, { target: "react" }),
      (error: unknown) =>
        error instanceof CodegenError &&
        error.diagnostics[0]?.code === "invalid-document" &&
        error.diagnostics[0]?.path === "deck/slide[s1]/table[t1]@columns[0]",
    );

    const badRowTextDeck = await expandDocument(
      createElement(
        Deck,
        null,
        createElement(
          Slide,
          { id: "s1" },
          createElement(
            Table,
            { id: "t1", columns: [100] },
            createElement(Row, { id: "r1" }, "stray row text"),
          ),
        ),
      ),
    );
    assert.throws(
      () => generate(badRowTextDeck, { target: "vanilla" }),
      (error: unknown) =>
        error instanceof CodegenError &&
        error.diagnostics[0]?.code === "invalid-document" &&
        error.diagnostics[0]?.path === "deck/slide[s1]/table[t1]/row[r1]",
    );

    const generateWithoutMotion = createCodegen(
      documentCapabilities.filter((mod) => mod.declaration !== "motion"),
    );
    const motionTableDeck = deck().append(
      slide({ id: "s1" }).append(
        table({ id: "t1", columns: [100], enter: fade() }).append(
          row().append(cell("ok")),
        ),
      ),
    );
    assert.throws(
      () => generateWithoutMotion(motionTableDeck, { target: "vanilla" }),
      (error: unknown) =>
        error instanceof CodegenError &&
        error.diagnostics[0]?.code === "unexpressible-value" &&
        error.diagnostics[0]?.path === "deck/slide[s1]/table[t1]@enter",
    );
  });

  it("generates vanilla and react code for imported ph-populated-placeholders.pptx slide 3 table placeholder", () => {
    const fixtureBytes = new Uint8Array(
      readFileSync(
        new URL(
          "../../../test/fixtures/ph-populated-placeholders.pptx",
          import.meta.url,
        ),
      ),
    );
    const importer = createImporter(documentCapabilities);
    const imported = importer(fixtureBytes);
    fixtureBytes.fill(0);

    const importedSlide3 = imported.document.children.find(
      (child): child is CommonNode =>
        child instanceof CommonNode &&
        child.type === "slide" &&
        child.id === "slide-3",
    );
    assert.ok(importedSlide3 !== undefined, "Expected slide-3 in fixture");

    // Isolate the table slide and clear the master background so this test focuses on table codegen.
    const importedMaster = imported.document.children.find(
      (child): child is CommonNode =>
        child instanceof CommonNode && child.type === "slide-master",
    );
    assert.ok(importedMaster !== undefined, "Expected slide-master in fixture");
    importedMaster.set({ background: undefined });
    for (const child of imported.document.children.filter(
      (item): item is CommonNode =>
        item instanceof CommonNode &&
        item.type === "slide" &&
        item !== importedSlide3,
    )) {
      child.remove();
    }

    const generate = createCodegen(documentCapabilities);
    const vanilla = generate(imported.document, { target: "vanilla" });
    const vanillaSource = vanilla.files[vanilla.entry] as string;
    assert.match(vanillaSource, /name: "Table Placeholder 2"/u);
    assert.match(vanillaSource, /placeholder: "tbl"/u);
    assert.doesNotMatch(vanillaSource, /nativePlaceholder/u);
    assert.match(vanillaSource, /columns: \[192, 192\]/u);

    const react = generate(imported.document, { target: "react" });
    const reactSource = react.files[react.entry] as string;
    assert.match(reactSource, /name="Table Placeholder 2"/u);
    assert.match(reactSource, /placeholder="tbl"/u);
    assert.match(reactSource, /columns=\{\[192, 192\]\}/u);
  });
});
