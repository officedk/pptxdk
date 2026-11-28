import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import { documentCapabilities } from "@pptxdk/capabilities/all";
import {
  cell,
  deck,
  font,
  row,
  slide,
  table,
} from "@pptxdk/capabilities/document";
import { tableCapability } from "@pptxdk/capabilities/table";
import { BuildError } from "@pptxdk/core";
import { createResolver } from "@pptxdk/core/build";
import {
  A_NS,
  OPC_REL_NS,
  P_NS,
  R_NS,
  REL_OFFICE_DOCUMENT,
  REL_SLIDE,
  createImporter,
} from "@pptxdk/import";
import { strToU8, zipSync } from "fflate";
import { describe, it } from "vitest";

const jostBook = new Uint8Array(
  readFileSync(
    new URL("../../../test/assets/fonts/Jost-400-Book.ttf", import.meta.url),
  ),
);
const jostBold = new Uint8Array(
  readFileSync(
    new URL("../../../test/assets/fonts/Jost-700-Bold.ttf", import.meta.url),
  ),
);
const measurementFont = font("Jost*", [
  { weight: 400, data: jostBook },
  { weight: 700, data: jostBold },
]);

const resolveTable = createResolver([tableCapability]);
const resolveAll = createResolver(documentCapabilities);
const importTable = createImporter([tableCapability]);

function tableDoc(node: ReturnType<typeof table>) {
  return deck({ width: 960, height: 540 }).append(
    slide({ id: "s1" }).append(node),
  );
}

function buildMinimalTablePptx(tblXml: string): Uint8Array {
  return zipSync({
    "[Content_Types].xml": strToU8(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
</Types>`,
    ),
    "_rels/.rels": strToU8(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${OPC_REL_NS}">
  <Relationship Id="rId1" Type="${REL_OFFICE_DOCUMENT}" Target="ppt/presentation.xml"/>
</Relationships>`,
    ),
    "ppt/presentation.xml": strToU8(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}">
  <p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst>
  <p:sldSz cx="9144000" cy="5143500"/>
  <p:notesSz cx="6858000" cy="9144000"/>
</p:presentation>`,
    ),
    "ppt/_rels/presentation.xml.rels": strToU8(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${OPC_REL_NS}">
  <Relationship Id="rId1" Type="${REL_SLIDE}" Target="slides/slide1.xml"/>
</Relationships>`,
    ),
    "ppt/slides/slide1.xml": strToU8(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr/>
      <p:graphicFrame>
        <p:nvGraphicFramePr><p:cNvPr id="2" name="Tbl"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>
        <p:xfrm><a:off x="190500" y="285750"/><a:ext cx="3429000" cy="741680"/></p:xfrm>
        <a:graphic>
          <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
            ${tblXml}
          </a:graphicData>
        </a:graphic>
      </p:graphicFrame>
    </p:spTree>
  </p:cSld>
</p:sld>`,
    ),
  });
}

describe("table public contract schemas", () => {
  it("validates table columns, spacing, vertical alignment, and table/row/cell styles via public resolve", () => {
    const valid = resolveAll(
      tableDoc(
        table({
          id: "pricing",
          columns: [120, "1.5fr", "auto"],
          style: {
            position: "absolute",
            left: 20,
            top: 30,
            width: 360,
            fontFamily: measurementFont,
            fontSize: "12pt",
            fontWeight: "bold",
            color: "#0F172A",
            align: "center",
            verticalAlign: "middle",
            border: "none",
            borderTop: { color: "#0F172A", width: 2 },
            borderBottom: { color: "#0F172A", width: 2 },
          },
        }).append(
          row({
            id: "r1",
            style: {
              height: 32,
              background: "#1E293B",
              color: "#FFFFFF",
              align: "right",
              borderBottom: { color: "#CBD5E1", width: 1 },
              borderLeft: "none",
            },
          }).append(
            cell("Header", {
              id: "c1",
              style: {
                background: "#EFF6FF",
                border: { color: "#2563EB", width: 2, style: "dashed" },
                padding: [6, 10],
              },
            }),
            cell("None", {
              style: {
                background: "none",
                border: "none",
                padding: [4, 8, 6, 10],
              },
            }),
            cell("Single", {
              style: {
                padding: 8,
                paddingLeft: 12,
              },
            }),
          ),
        ),
      ),
    );
    assert.equal(valid.slides.length, 1);

    for (const [label, buildInvalid] of [
      [
        "empty columns",
        () =>
          tableDoc(
            table({
              columns: [],
              style: { position: "absolute", left: 20, top: 30, width: 300 },
            }).append(row().append(cell("A"))),
          ),
      ],
      [
        "zero fr column",
        () =>
          tableDoc(
            table({
              columns: ["0fr"],
              style: { position: "absolute", left: 20, top: 30, width: 300 },
            }).append(row().append(cell("A"))),
          ),
      ],
      [
        "percentage column",
        () =>
          tableDoc(
            table({
              columns: ["50%" as never],
              style: { position: "absolute", left: 20, top: 30, width: 300 },
            }).append(row().append(cell("A"))),
          ),
      ],
      [
        "invalid verticalAlign",
        () =>
          tableDoc(
            table({
              columns: [120],
              style: {
                position: "absolute",
                left: 20,
                top: 30,
                width: 120,
                verticalAlign: "center" as never,
              },
            }).append(row().append(cell("A"))),
          ),
      ],
      [
        "invalid align on Table",
        () =>
          tableDoc(
            table({
              columns: [120],
              style: {
                position: "absolute",
                left: 20,
                top: 30,
                width: 120,
                align: "middle" as never,
              },
            }).append(row().append(cell("A", { style: { align: "left" } }))),
          ),
      ],
      [
        "invalid borderTop on Table",
        () =>
          tableDoc(
            table({
              columns: [120],
              style: {
                position: "absolute",
                left: 20,
                top: 30,
                width: 120,
                borderTop: { color: "#0F172A", style: "double" } as never,
              },
            }).append(row().append(cell("A"))),
          ),
      ],
      [
        "invalid borderBottom on Row",
        () =>
          tableDoc(
            table({
              columns: [120],
              style: { position: "absolute", left: 20, top: 30, width: 120 },
            }).append(
              row({
                style: {
                  borderBottom: { color: "#0F172A", width: -1 } as never,
                },
              }).append(cell("A")),
            ),
          ),
      ],
      [
        "zero row height",
        () =>
          tableDoc(
            table({
              columns: [120],
              style: { position: "absolute", left: 20, top: 30, width: 120 },
            }).append(row({ style: { height: 0 } }).append(cell("A"))),
          ),
      ],
      [
        "3-element padding tuple",
        () =>
          tableDoc(
            table({
              columns: [120],
              style: { position: "absolute", left: 20, top: 30, width: 120 },
            }).append(
              row().append(
                cell("A", { style: { padding: [4, 8, 6] as never } }),
              ),
            ),
          ),
      ],
      [
        "unsupported double border style",
        () =>
          tableDoc(
            table({
              columns: [120],
              style: { position: "absolute", left: 20, top: 30, width: 120 },
            }).append(
              row().append(
                cell("A", {
                  style: {
                    border: { color: "#0F172A", style: "double" } as never,
                  },
                }),
              ),
            ),
          ),
      ],
      [
        "zero colSpan",
        () =>
          tableDoc(
            table({
              columns: [120],
              style: { position: "absolute", left: 20, top: 30, width: 120 },
            }).append(row().append(cell("A", { colSpan: 0 }))),
          ),
      ],
      [
        "fractional rowSpan",
        () =>
          tableDoc(
            table({
              columns: [120],
              style: { position: "absolute", left: 20, top: 30, width: 120 },
            }).append(row().append(cell("A", { rowSpan: 1.5 }))),
          ),
      ],
    ] as const) {
      assert.throws(
        () => resolveTable(buildInvalid()),
        (error: unknown) =>
          error instanceof BuildError && error.diagnostics.length > 0,
        `${label} should throw BuildError`,
      );
    }
  });

  it("imports valid OOXML <a:tbl> structures and rejects malformed grid columns gracefully", () => {
    const validPptx = buildMinimalTablePptx(`
      <a:tbl>
        <a:tblPr firstRow="1" bandRow="1" rtl="0"/>
        <a:tblGrid><a:gridCol w="1714500"/><a:gridCol w="1714500"/></a:tblGrid>
        <a:tr h="370840">
          <a:tc gridSpan="2"><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Merged</a:t></a:r></a:p></a:txBody><a:tcPr marL="95250" marT="57150" anchor="ctr"/></a:tc>
          <a:tc hMerge="1"><a:txBody><a:bodyPr/><a:lstStyle/><a:p/></a:txBody><a:tcPr/></a:tc>
        </a:tr>
      </a:tbl>
    `);
    const validResult = importTable(validPptx);
    assert.deepEqual(validResult.report.diagnostics, []);
    const slideNode = validResult.document.children[0];
    const importedTable =
      slideNode !== undefined && "children" in slideNode
        ? slideNode.children.find((node) => node.type === "table")
        : undefined;
    assert.ok(importedTable !== undefined);

    const invalidGridPptx = buildMinimalTablePptx(`
      <a:tbl>
        <a:tblPr/>
        <a:tblGrid><a:gridCol w="0"/></a:tblGrid>
        <a:tr h="370840"><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p/></a:txBody><a:tcPr/></a:tc></a:tr>
      </a:tbl>
    `);
    const invalidResult = importTable(invalidGridPptx);
    assert.ok(invalidResult.report.diagnostics.length > 0);
  });
});
