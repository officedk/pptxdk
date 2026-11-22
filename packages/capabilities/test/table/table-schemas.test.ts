import assert from "node:assert/strict";

import { describe, it } from "vitest";

import {
  cellDeclarationPropsSchema,
  rowDeclarationPropsSchema,
  tableCellStyleSchema,
  tableColumnSchema,
  tableDeclarationPropsSchema,
  tableRowStyleSchema,
  tableSpacingSchema,
  tableStyleSchema,
  tableTextStyleSchema,
  tableVerticalAlignSchema,
  tblCellPrXmlSchema,
  tblCellXmlSchema,
  tblGridColXmlSchema,
  tblPrXmlSchema,
  tblRowXmlSchema,
} from "../../src/table/table-schemas.ts";

describe("table Zod v4 schemas", () => {
  it("validates table columns, spacing, vertical alignment, and styles", () => {
    assert.equal(tableColumnSchema.safeParse(120).success, true);
    assert.equal(tableColumnSchema.safeParse("1.5fr").success, true);
    assert.equal(tableColumnSchema.safeParse("auto").success, true);
    assert.equal(tableColumnSchema.safeParse(0).success, false);
    assert.equal(tableColumnSchema.safeParse("0fr").success, false);
    assert.equal(tableColumnSchema.safeParse("-10fr").success, false);
    assert.equal(tableColumnSchema.safeParse("50%").success, false);

    assert.equal(tableVerticalAlignSchema.safeParse("top").success, true);
    assert.equal(tableVerticalAlignSchema.safeParse("middle").success, true);
    assert.equal(tableVerticalAlignSchema.safeParse("bottom").success, true);
    assert.equal(tableVerticalAlignSchema.safeParse("center").success, false);

    assert.equal(tableSpacingSchema.safeParse(8).success, true);
    assert.equal(tableSpacingSchema.safeParse([4, 8]).success, true);
    assert.equal(tableSpacingSchema.safeParse([4, 8, 6, 10]).success, true);
    assert.equal(tableSpacingSchema.safeParse([4, 8, 6]).success, false);
    assert.equal(tableSpacingSchema.safeParse(-1).success, false);

    assert.equal(
      tableTextStyleSchema.safeParse({
        fontFamily: "Aptos",
        fontSize: "12pt",
        fontWeight: "bold",
        color: "#0F172A",
        verticalAlign: "middle",
      }).success,
      true,
    );
    assert.equal(
      tableStyleSchema.safeParse({
        fontSize: 14,
        flexGrow: 1,
        width: 400,
      }).success,
      true,
    );
    assert.equal(
      tableRowStyleSchema.safeParse({
        height: 32,
        backgroundColor: "#1E293B",
        color: "#FFFFFF",
      }).success,
      true,
    );
    assert.equal(tableRowStyleSchema.safeParse({ height: 0 }).success, false);
    assert.equal(
      tableCellStyleSchema.safeParse({
        backgroundColor: "#EFF6FF",
        borderColor: "#2563EB",
        borderWidth: 2,
        borderStyle: "dashed",
        padding: [6, 10],
      }).success,
      true,
    );
    assert.equal(
      tableCellStyleSchema.safeParse({ borderStyle: "double" }).success,
      false,
    );
  });

  it("validates declaration props for Table, Row, and Cell", () => {
    assert.equal(
      tableDeclarationPropsSchema.safeParse({
        id: "pricing",
        x: 20,
        y: 30,
        width: 360,
        columns: [120, "1fr", "auto"],
      }).success,
      true,
    );
    assert.equal(
      tableDeclarationPropsSchema.safeParse({ columns: [] }).success,
      false,
    );
    assert.equal(
      rowDeclarationPropsSchema.safeParse({
        id: "r1",
        style: { height: 28 },
      }).success,
      true,
    );
    assert.equal(
      cellDeclarationPropsSchema.safeParse({
        id: "c1",
        colSpan: 2,
        rowSpan: 2,
        style: { paddingLeft: 8 },
      }).success,
      true,
    );
    assert.equal(
      cellDeclarationPropsSchema.safeParse({ colSpan: 0 }).success,
      false,
    );
    assert.equal(
      cellDeclarationPropsSchema.safeParse({ rowSpan: 1.5 }).success,
      false,
    );
  });

  it("validates OOXML <a:tbl> fragment schemas", () => {
    assert.deepEqual(
      tblPrXmlSchema.parse({ firstRow: "1", bandRow: "1", rtl: "0" }),
      { firstRow: true, bandRow: true, rtl: false },
    );
    assert.deepEqual(tblGridColXmlSchema.parse({ w: "1828800" }), {
      w: 1_828_800,
    });
    assert.equal(tblGridColXmlSchema.safeParse({ w: "0" }).success, false);
    assert.deepEqual(tblRowXmlSchema.parse({ h: "370840" }), { h: 370_840 });
    assert.deepEqual(
      tblCellXmlSchema.parse({
        gridSpan: "2",
        rowSpan: "2",
        hMerge: "1",
        vMerge: "0",
      }),
      { gridSpan: 2, rowSpan: 2, hMerge: true, vMerge: false },
    );
    assert.deepEqual(
      tblCellPrXmlSchema.parse({
        marL: "95250",
        marT: "57150",
        anchor: "ctr",
      }),
      { marL: 95_250, marT: 57_150, anchor: "ctr" },
    );
  });
});
